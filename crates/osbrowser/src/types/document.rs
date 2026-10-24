//! `Document`: webcore's document API over the system browser's live DOM.
//!
//! Node ids follow webcore's conventions: a `u32`, `0` meaning "no node", and
//! [`DocumentRoot::node_id`] naming `<html>`. Writes are queued in order and not
//! waited for; reads are one round trip that also delivers every write queued
//! before them. A node's tag, local name, namespace and type never change, so
//! they are asked for once and cached.

use std::cell::RefCell;
use std::collections::{HashMap, HashSet};
use std::sync::{Mutex, OnceLock};

use serde::Deserialize;
use serde_json::{Value, json};

use super::{DocumentKind, FormEvent, FormEventCallback, FormEventKind, PendingImagePoll, Rect};
use crate::dom::events::{DomEvent, EventHandler, ListenerOptions};
use crate::dom::HtmlEventType;
use crate::page;

/// The document element, as webcore's `Document::root` exposes it.
#[derive(Clone, Copy, Debug, Default)]
pub struct DocumentRoot {
    pub node_id: u32,
}

struct Listener {
    node: u32,
    handler: Option<EventHandler>,
    once: bool,
}

/// What never changes about a node, fetched once.
#[derive(Clone, Copy, Default)]
struct Facts {
    tag: Option<&'static str>,
    local: &'static str,
    name: &'static str,
    namespace: Option<&'static str>,
    prefix: Option<&'static str>,
    node_type: u16,
}

#[derive(Deserialize, Default)]
struct FactsWire {
    tag: Option<String>,
    local: String,
    name: String,
    namespace: Option<String>,
    prefix: Option<String>,
    node_type: u16,
}

/// Names are a small, closed set — tags, attribute-free local names — so they
/// are interned for the life of the process and handed out as `&str`.
fn intern(text: &str) -> &'static str {
    static NAMES: OnceLock<Mutex<HashSet<&'static str>>> = OnceLock::new();
    let mut names = NAMES.get_or_init(Default::default).lock().unwrap();
    if let Some(name) = names.get(text) {
        return name;
    }
    let name: &'static str = Box::leak(text.to_owned().into_boxed_str());
    names.insert(name);
    name
}

fn some(id: u32) -> Option<u32> {
    (id != 0).then_some(id)
}

pub struct Document {
    pub root: DocumentRoot,
    pub title: String,
    pub kind: DocumentKind,
    pub scroll_x: f32,
    pub scroll_y: f32,
    pub viewport_w: f32,
    pub viewport_h: f32,
    pub on_form_event: Option<FormEventCallback>,
    page: u64,
    listeners: HashMap<u32, Listener>,
    next_listener: u32,
    facts: RefCell<HashMap<u32, Facts>>,
}

impl Document {
    /// Open a page in the browser and load `html` into it.
    pub(crate) fn load(html: &str, viewport_w: f32, viewport_h: f32) -> Self {
        let page = page::open_page("");
        let root = page::read::<u32>(page, "load_html", json!([html]));
        let title = page::read::<String>(page, "title", json!([]));
        Self {
            root: DocumentRoot { node_id: root },
            title,
            kind: DocumentKind::Html,
            scroll_x: 0.0,
            scroll_y: 0.0,
            viewport_w,
            viewport_h,
            on_form_event: None,
            page,
            listeners: HashMap::new(),
            next_listener: 1,
            facts: RefCell::new(HashMap::new()),
        }
    }

    fn get<T: serde::de::DeserializeOwned + Default>(&self, operation: &str, args: Value) -> T {
        page::read(self.page, operation, args)
    }

    fn put(&self, operation: &str, args: Value) {
        page::write(self.page, operation, args);
    }

    fn make(&self, operation: &str, args: Value) -> u32 {
        page::create(self.page, operation, args)
    }

    fn facts(&self, id: u32) -> Facts {
        if id == 0 {
            return Facts::default();
        }
        if let Some(facts) = self.facts.borrow().get(&id) {
            return *facts;
        }
        let wire: FactsWire = self.get("facts", json!([id]));
        let facts = Facts {
            tag: wire.tag.as_deref().map(intern),
            local: intern(&wire.local),
            name: intern(&wire.name),
            namespace: wire.namespace.as_deref().map(intern),
            prefix: wire.prefix.as_deref().map(intern),
            node_type: wire.node_type,
        };
        if wire.node_type != 0 {
            self.facts.borrow_mut().insert(id, facts);
        }
        facts
    }

    /// Run the listeners for every event the browser reported for this page.
    ///
    /// webcore runs a listener inside the call that caused it; here the event
    /// happens in the browser, so the handler runs when the report arrives —
    /// in the window loop, and before any dispatching call returns.
    pub(crate) fn pump(&mut self) {
        for event in page::take_events(self.page) {
            match event.listener {
                Some(listener) => self.run_listener(listener, &event),
                None => self.observe(&event),
            }
        }
    }

    fn run_listener(&mut self, listener: u32, event: &crate::BrowserEvent) {
        let Some(entry) = self.listeners.get_mut(&listener) else { return };
        let Some(mut handler) = entry.handler.take() else { return };
        let (node, once) = (entry.node, entry.once);
        let mut dom_event = DomEvent::from_browser(&event.kind, event.node as u32, node, &event.fields);
        handler(&mut dom_event, self);
        if once {
            self.listeners.remove(&listener);
        } else if let Some(entry) = self.listeners.get_mut(&listener) {
            entry.handler = Some(handler);
        }
    }

    /// The page-wide report: viewport, scroll and form interaction.
    fn observe(&mut self, event: &crate::BrowserEvent) {
        let f = &event.fields;
        let number = |key: &str| f.get(key).and_then(Value::as_f64).unwrap_or(0.0) as f32;
        match event.kind.as_str() {
            "resize" => {
                self.viewport_w = number("width");
                self.viewport_h = number("height");
            }
            "scroll" => {
                self.scroll_x = number("scroll_x");
                self.scroll_y = number("scroll_y");
            }
            _ => {}
        }
        let Some(form) = f.get("form") else { return };
        let text = |key: &str| form.get(key).and_then(Value::as_str).unwrap_or("").to_string();
        let checkable = matches!(text("type").as_str(), "checkbox" | "radio");
        let kind = match event.kind.as_str() {
            "input" if !checkable => FormEventKind::Input(text("value")),
            "change" if !checkable => FormEventKind::Change(text("value")),
            "click" if checkable => {
                FormEventKind::Toggle(form.get("checked").and_then(Value::as_bool).unwrap_or(false))
            }
            "click" if matches!(text("tag").as_str(), "button")
                || matches!(text("type").as_str(), "button" | "submit" | "reset") =>
            {
                FormEventKind::Click(text("value"))
            }
            "submit" => FormEventKind::Submit(text("action")),
            "focus" => FormEventKind::Focus,
            "blur" => FormEventKind::Blur,
            _ => return,
        };
        let form_event = FormEvent {
            tag: text("tag"),
            id: text("id"),
            name: text("name"),
            kind,
            element: event.node as u32,
        };
        if let Some(callback) = self.on_form_event.as_mut() {
            callback(&form_event);
        }
    }

    // ── Creation ──

    pub fn create_element(&mut self, tag: &str) -> u32 {
        self.make("create_element", json!([tag]))
    }

    pub fn create_text_node(&mut self, text: &str) -> u32 {
        self.make("create_text_node", json!([text]))
    }

    pub fn create_comment(&mut self, text: &str) -> u32 {
        self.make("create_comment", json!([text]))
    }

    pub fn create_cdata_section(&mut self, data: &str) -> u32 {
        self.make("create_cdata_section", json!([data]))
    }

    pub fn create_element_ns(&mut self, namespace: &str, qualified_name: &str) -> u32 {
        self.make("create_element_ns", json!([namespace, qualified_name]))
    }

    pub fn create_processing_instruction(&mut self, target: &str, data: &str) -> u32 {
        self.make("create_processing_instruction", json!([target, data]))
    }

    pub fn create_document_fragment(&mut self) -> u32 {
        self.make("create_document_fragment", json!([]))
    }

    pub fn clone_node(&mut self, id: u32, deep: bool) -> u32 {
        if id == 0 {
            return 0;
        }
        self.make("clone_node", json!([id, deep]))
    }

    // ── Document-level ──

    pub fn title(&self) -> String {
        self.get("title", json!([]))
    }

    pub fn set_title(&mut self, title: &str) {
        self.title = title.to_string();
        self.put("set_title", json!([title]));
    }

    pub fn kind(&self) -> DocumentKind {
        self.kind
    }

    pub fn body(&self) -> Option<u32> {
        some(self.get("body", json!([])))
    }

    pub fn head(&self) -> Option<u32> {
        some(self.get("head", json!([])))
    }

    pub fn document_element(&self) -> Option<u32> {
        some(self.root.node_id).or_else(|| some(self.get("document_element", json!([]))))
    }

    pub fn set_viewport(&mut self, width: f32, height: f32) {
        self.viewport_w = width;
        self.viewport_h = height;
    }

    /// The browser lays out on demand; there is nothing to flush from here.
    pub fn flush_layout(&mut self) {}

    /// The browser fetches and decodes the page's images itself.
    pub fn poll_pending_images_budgeted(
        &mut self,
        _max_images: usize,
        _max_time: std::time::Duration,
    ) -> PendingImagePoll {
        PendingImagePoll::default()
    }

    // ── Queries ──

    pub fn get_element_by_id(&self, id: &str) -> Option<u32> {
        some(self.get("get_element_by_id", json!([id])))
    }

    pub fn get_elements_by_tag_name(&self, tag: &str) -> Vec<u32> {
        self.get("get_elements_by_tag_name", json!([tag]))
    }

    pub fn get_elements_by_class_name(&self, names: &str) -> Vec<u32> {
        self.get("get_elements_by_class_name", json!([names]))
    }

    pub fn query_selector(&self, selector: &str) -> Option<u32> {
        some(self.get("query_selector", json!([selector])))
    }

    pub fn query_selector_all(&self, selector: &str) -> Vec<u32> {
        self.get("query_selector_all", json!([selector]))
    }

    pub fn query_selector_within(&self, root: u32, selector: &str) -> Option<u32> {
        some(self.get("query_selector_within", json!([root, selector])))
    }

    pub fn query_selector_all_within(&self, root: u32, selector: &str) -> Vec<u32> {
        self.get("query_selector_all_within", json!([root, selector]))
    }

    pub fn matches(&self, id: u32, selectors: &str) -> bool {
        self.get("matches", json!([id, selectors]))
    }

    pub fn closest(&self, id: u32, selectors: &str) -> Option<u32> {
        some(self.get("closest", json!([id, selectors])))
    }

    pub fn element_from_point(&self, x: f32, y: f32) -> Option<u32> {
        some(self.get("element_from_point", json!([x, y])))
    }

    // ── Tree ──

    pub fn append_child(&mut self, parent_id: u32, child_id: u32) {
        if parent_id != 0 && child_id != 0 {
            self.put("append_child", json!([parent_id, child_id]));
        }
    }

    pub fn insert_before(&mut self, parent_id: u32, child_id: u32, reference_id: u32) {
        if parent_id != 0 && child_id != 0 && reference_id != 0 {
            self.put("insert_before", json!([parent_id, child_id, reference_id]));
        }
    }

    pub fn remove_child(&mut self, child_id: u32) {
        if child_id != 0 {
            self.put("remove_child", json!([child_id]));
        }
    }

    pub fn replace_child(&mut self, parent: u32, new_child: u32, old_child: u32) -> bool {
        self.get("replace_child", json!([parent, new_child, old_child]))
    }

    pub fn append(&mut self, parent: u32, nodes: &[u32]) {
        self.put("append", json!([parent, nodes]));
    }

    pub fn prepend(&mut self, parent: u32, nodes: &[u32]) {
        self.put("prepend", json!([parent, nodes]));
    }

    pub fn before(&mut self, id: u32, nodes: &[u32]) {
        self.put("before", json!([id, nodes]));
    }

    pub fn after(&mut self, id: u32, nodes: &[u32]) {
        self.put("after", json!([id, nodes]));
    }

    pub fn replace_with(&mut self, id: u32, nodes: &[u32]) {
        self.put("replace_with", json!([id, nodes]));
    }

    pub fn insert_adjacent_element(&mut self, id: u32, position: &str, other: u32) -> Option<u32> {
        some(self.get("insert_adjacent_element", json!([id, position, other])))
    }

    pub fn insert_adjacent_text(&mut self, id: u32, position: &str, data: &str) {
        self.put("insert_adjacent_text", json!([id, position, data]));
    }

    pub fn normalize(&mut self, id: u32) {
        self.put("normalize", json!([id]));
    }

    pub fn parent_node(&self, id: u32) -> u32 {
        self.get("parent_node", json!([id]))
    }

    pub fn parent_element(&self, id: u32) -> Option<u32> {
        some(self.get("parent_element", json!([id])))
    }

    pub fn child_nodes(&self, id: u32) -> Vec<u32> {
        self.get("child_nodes", json!([id]))
    }

    pub fn children(&self, id: u32) -> Vec<u32> {
        self.get("children", json!([id]))
    }

    pub fn first_child(&self, id: u32) -> Option<u32> {
        some(self.get("first_child", json!([id])))
    }

    pub fn last_child(&self, id: u32) -> Option<u32> {
        some(self.get("last_child", json!([id])))
    }

    pub fn next_sibling(&self, id: u32) -> u32 {
        self.get("next_sibling", json!([id]))
    }

    pub fn previous_sibling(&self, id: u32) -> u32 {
        self.get("previous_sibling", json!([id]))
    }

    pub fn first_element_child(&self, id: u32) -> Option<u32> {
        some(self.get("first_element_child", json!([id])))
    }

    pub fn last_element_child(&self, id: u32) -> Option<u32> {
        some(self.get("last_element_child", json!([id])))
    }

    pub fn next_element_sibling(&self, id: u32) -> Option<u32> {
        some(self.get("next_element_sibling", json!([id])))
    }

    pub fn previous_element_sibling(&self, id: u32) -> Option<u32> {
        some(self.get("previous_element_sibling", json!([id])))
    }

    pub fn child_element_count(&self, id: u32) -> usize {
        self.get("child_element_count", json!([id]))
    }

    pub fn has_child_nodes(&self, id: u32) -> bool {
        self.get("has_child_nodes", json!([id]))
    }

    pub fn contains(&self, id: u32, other: u32) -> bool {
        self.get("contains", json!([id, other]))
    }

    pub fn is_connected(&self, id: u32) -> bool {
        self.get("is_connected", json!([id]))
    }

    pub fn is_equal_node(&self, a: u32, b: u32) -> bool {
        self.get("is_equal_node", json!([a, b]))
    }

    pub fn compare_document_position(&self, a: u32, b: u32) -> u16 {
        self.get("compare_document_position", json!([a, b]))
    }

    // ── What a node is ──

    pub fn node_type(&self, id: u32) -> u16 {
        self.facts(id).node_type
    }

    pub fn node_name(&self, id: u32) -> String {
        self.facts(id).name.to_string()
    }

    pub fn tag_name(&self, id: u32) -> Option<&str> {
        self.facts(id).tag
    }

    pub fn local_name(&self, id: u32) -> String {
        self.facts(id).local.to_string()
    }

    pub fn namespace_uri(&self, id: u32) -> Option<String> {
        self.facts(id).namespace.map(str::to_string)
    }

    pub fn prefix(&self, id: u32) -> Option<String> {
        self.facts(id).prefix.map(str::to_string)
    }

    pub fn is_element(&self, id: u32) -> bool {
        self.node_type(id) == 1
    }

    pub fn is_text_node(&self, id: u32) -> bool {
        self.node_type(id) == 3
    }

    pub fn is_comment_node(&self, id: u32) -> bool {
        self.node_type(id) == 8
    }

    pub fn is_character_data(&self, id: u32) -> bool {
        matches!(self.node_type(id), 3 | 4 | 7 | 8)
    }

    pub fn is_document_fragment(&self, id: u32) -> bool {
        self.node_type(id) == 11
    }

    // ── Content ──

    pub fn node_value(&self, id: u32) -> Option<String> {
        self.get("node_value", json!([id]))
    }

    pub fn text_content(&self, id: u32) -> String {
        self.get("text_content", json!([id]))
    }

    pub fn set_text_content(&mut self, id: u32, text: &str) {
        self.put("set_text_content", json!([id, text]));
    }

    pub fn text_data(&self, id: u32) -> String {
        self.get("text_data", json!([id]))
    }

    pub fn set_text_data(&mut self, id: u32, data: &str) {
        self.put("set_text_data", json!([id, data]));
    }

    pub fn inner_html(&self, id: u32) -> String {
        self.get("inner_html", json!([id]))
    }

    pub fn outer_html(&self, id: u32) -> String {
        self.get("outer_html", json!([id]))
    }

    pub fn inner_text(&self, id: u32) -> String {
        self.get("inner_text", json!([id]))
    }

    // ── Attributes ──

    pub fn get_attribute(&self, id: u32, key: &str) -> Option<String> {
        self.get("get_attribute", json!([id, key]))
    }

    pub fn set_attribute(&mut self, id: u32, key: &str, value: &str) {
        self.put("set_attribute", json!([id, key, value]));
    }

    pub fn remove_attribute(&mut self, id: u32, key: &str) {
        self.put("remove_attribute", json!([id, key]));
    }

    pub fn has_attribute(&self, id: u32, name: &str) -> bool {
        self.get("has_attribute", json!([id, name]))
    }

    pub fn has_attributes(&self, id: u32) -> bool {
        self.get("has_attributes", json!([id]))
    }

    pub fn toggle_attribute(&mut self, id: u32, name: &str) -> bool {
        self.get("toggle_attribute", json!([id, name]))
    }

    pub fn get_attribute_names(&self, id: u32) -> Vec<String> {
        self.get("get_attribute_names", json!([id]))
    }

    pub fn get_attribute_ns(&self, id: u32, namespace: &str, local_name: &str) -> Option<String> {
        self.get("get_attribute_ns", json!([id, namespace, local_name]))
    }

    pub fn set_attribute_ns(
        &mut self,
        id: u32,
        namespace: &str,
        qualified_name: &str,
        value: &str,
    ) {
        self.put("set_attribute_ns", json!([id, namespace, qualified_name, value]));
    }

    pub fn has_attribute_ns(&self, id: u32, namespace: &str, local_name: &str) -> bool {
        self.get("has_attribute_ns", json!([id, namespace, local_name]))
    }

    pub fn remove_attribute_ns(&mut self, id: u32, namespace: &str, local_name: &str) {
        self.put("remove_attribute_ns", json!([id, namespace, local_name]));
    }

    pub fn id(&self, id: u32) -> String {
        self.get("id", json!([id]))
    }

    pub fn set_id(&mut self, id: u32, value: &str) {
        self.put("set_id", json!([id, value]));
    }

    pub fn class_name(&self, id: u32) -> String {
        self.get("class_name", json!([id]))
    }

    pub fn set_class_name(&mut self, id: u32, value: &str) {
        self.put("set_class_name", json!([id, value]));
    }

    pub fn class_list_add(&mut self, id: u32, class: &str) {
        self.put("class_list_add", json!([id, class]));
    }

    pub fn class_list_remove(&mut self, id: u32, class: &str) {
        self.put("class_list_remove", json!([id, class]));
    }

    pub fn class_list_contains(&self, id: u32, class: &str) -> bool {
        self.get("class_list_contains", json!([id, class]))
    }

    pub fn class_list_toggle(&mut self, id: u32, class: &str) -> bool {
        self.get("class_list_toggle", json!([id, class]))
    }

    pub fn dataset(&self, id: u32) -> Vec<(String, String)> {
        self.get("dataset", json!([id]))
    }

    pub fn dataset_get(&self, id: u32, key: &str) -> Option<String> {
        self.get("dataset_get", json!([id, key]))
    }

    pub fn set_dataset(&mut self, id: u32, key: &str, value: &str) {
        self.put("set_dataset", json!([id, key, value]));
    }

    pub fn remove_dataset(&mut self, id: u32, key: &str) {
        self.put("remove_dataset", json!([id, key]));
    }

    pub fn tab_index(&self, id: u32) -> i32 {
        self.get("tab_index", json!([id]))
    }

    pub fn set_tab_index(&mut self, id: u32, index: i32) {
        self.put("set_tab_index", json!([id, index]));
    }

    // ── Style ──

    pub fn get_style_property(&self, id: u32, prop: &str) -> Option<String> {
        self.get("get_style_property", json!([id, prop]))
    }

    pub fn set_style_property(&mut self, id: u32, prop: &str, value: &str) {
        self.put("set_style_property", json!([id, prop, value]));
    }

    pub fn remove_style_property(&mut self, id: u32, prop: &str) -> Option<String> {
        self.get("remove_style_property", json!([id, prop]))
    }

    pub fn style_property_len(&self, id: u32) -> usize {
        self.get("style_property_len", json!([id]))
    }

    pub fn style_property_item(&self, id: u32, index: usize) -> Option<String> {
        self.get("style_property_item", json!([id, index]))
    }

    pub fn computed_style_property(&mut self, id: u32, property: &str) -> String {
        self.get("computed_style_property", json!([id, property]))
    }

    // ── Geometry (CSSOM View) ──

    pub fn get_bounding_client_rect(&self, id: u32) -> Option<Rect> {
        self.get("get_bounding_client_rect", json!([id]))
    }

    pub fn get_client_rects(&self, id: u32) -> Vec<Rect> {
        self.get("get_client_rects", json!([id]))
    }

    pub fn offset_width(&self, id: u32) -> f32 {
        self.get("offset_width", json!([id]))
    }

    pub fn offset_height(&self, id: u32) -> f32 {
        self.get("offset_height", json!([id]))
    }

    pub fn offset_top(&mut self, id: u32) -> f32 {
        self.get("offset_top", json!([id]))
    }

    pub fn offset_left(&mut self, id: u32) -> f32 {
        self.get("offset_left", json!([id]))
    }

    pub fn offset_parent(&mut self, id: u32) -> Option<u32> {
        some(self.get("offset_parent", json!([id])))
    }

    pub fn scroll_into_view(&mut self, id: u32) {
        self.put("scroll_into_view", json!([id]));
    }

    // ── Form controls ──

    pub fn value(&self, id: u32) -> String {
        self.get("value", json!([id]))
    }

    pub fn set_value(&mut self, id: u32, value: &str) {
        self.put("set_value", json!([id, value]));
    }

    pub fn checked(&self, id: u32) -> bool {
        self.get("checked", json!([id]))
    }

    pub fn set_checked(&mut self, id: u32, checked: bool) {
        self.put("set_checked", json!([id, checked]));
    }

    pub fn add_item(&mut self, select: u32, text: &str) {
        self.put("add_item", json!([select, text]));
    }

    pub fn remove_item(&mut self, select: u32, index: usize) {
        self.put("remove_item", json!([select, index]));
    }

    pub fn clear_items(&mut self, select: u32) {
        self.put("clear_items", json!([select]));
    }

    pub fn item_text(&self, select: u32, index: usize) -> String {
        self.get("item_text", json!([select, index]))
    }

    pub fn set_item_text(&mut self, select: u32, index: usize, text: &str) {
        self.put("set_item_text", json!([select, index, text]));
    }

    pub fn selected_index(&self, select: u32) -> i32 {
        self.get("selected_index", json!([select]))
    }

    pub fn set_selected_index(&mut self, select: u32, index: i32) {
        self.put("set_selected_index", json!([select, index]));
    }

    pub fn focus(&mut self, id: u32) {
        self.put("focus", json!([id]));
    }

    pub fn click(&mut self, id: u32) {
        self.get::<bool>("click", json!([id]));
        self.pump();
    }

    // ── HTMLDialogElement ──

    pub fn show_dialog(&mut self, id: u32, modal: bool) {
        self.put("show_dialog", json!([id, modal]));
    }

    pub fn close_dialog(&mut self, id: u32) {
        self.put("close_dialog", json!([id]));
    }

    pub fn dialog_open(&self, id: u32) -> bool {
        self.get("dialog_open", json!([id]))
    }

    // ── Events ──

    pub fn add_event_listener(
        &mut self,
        id: u32,
        event_type: &str,
        handler: EventHandler,
        options: ListenerOptions,
    ) -> u32 {
        let listener = self.next_listener;
        self.next_listener += 1;
        self.listeners.insert(listener, Listener { node: id, handler: Some(handler), once: options.once });
        self.put(
            "add_event_listener",
            json!([id, event_type, listener, options.capture, options.once, options.passive]),
        );
        listener
    }

    pub fn remove_event_listener(&mut self, listener_id: u32) {
        if self.listeners.remove(&listener_id).is_some() {
            self.put("remove_event_listener", json!([listener_id]));
        }
    }

    /// Deliver a pointer event at a document point, the way a real one would
    /// arrive: hit-tested, with `click` following a matched press and release.
    pub fn process_mouse_event(
        &mut self,
        etype: HtmlEventType,
        doc_pt: (f32, f32),
        button: u8,
    ) -> bool {
        let changed = self.get("process_mouse_event", json!([etype.as_str(), doc_pt.0, doc_pt.1, button]));
        self.pump();
        changed
    }

    /// Move the hover to the element at `doc_pt`, firing the over/out pairs.
    pub fn dispatch_over_out(&mut self, doc_pt: (f32, f32)) -> bool {
        let changed = self.get("dispatch_over_out", json!([doc_pt.0, doc_pt.1]));
        self.pump();
        changed
    }

    pub fn process_keyboard_event(&mut self, event: crate::dom::events::DomEvent) -> bool {
        let handled = self.get(
            "process_keyboard_event",
            json!([{
                "kind": event.event_type,
                "key": event.key(),
                "code": event.code(),
                "key_code": event.key_code,
                "repeat": event.repeat(),
                "ctrl_key": event.ctrl_key,
                "shift_key": event.shift_key,
                "alt_key": event.alt_key,
                "meta_key": event.meta_key,
            }]),
        );
        self.pump();
        handled
    }

    /// Scroll the element under `doc_pt`. A positive `delta_y` scrolls up, as
    /// webcore's does.
    pub fn process_wheel_event_xy(
        &mut self,
        doc_pt: (f32, f32),
        delta_x: f32,
        delta_y: f32,
    ) -> bool {
        let changed = self.get("process_wheel_event_xy", json!([doc_pt.0, doc_pt.1, delta_x, delta_y]));
        self.pump();
        changed
    }

    pub fn dispatch_input_event(
        &mut self,
        mut evt: crate::dom::HtmlEvent,
    ) -> (bool, crate::dom::HtmlEvent) {
        #[derive(Deserialize, Default)]
        struct Outcome {
            dispatched: bool,
            prevented: bool,
        }
        let outcome: Outcome = self.get(
            "dispatch_input_event",
            json!([{
                "kind": evt.event_type.as_str(),
                "target": evt.target,
                "client_x": evt.client_pos.0,
                "client_y": evt.client_pos.1,
                "button": evt.button,
                "delta_x": evt.delta_x,
                "delta_y": evt.delta_y,
                "ctrl_key": evt.ctrl_key,
                "shift_key": evt.shift_key,
                "alt_key": evt.alt_key,
                "meta_key": evt.meta_key,
            }]),
        );
        self.pump();
        evt.default_prevented = outcome.prevented;
        (outcome.dispatched, evt)
    }

    // ── Canvas ──

    /// Give a `<canvas>` its 2D context. Answers whether the element is one.
    pub fn get_context_2d(&mut self, id: u32) -> bool {
        self.local_name(id).eq_ignore_ascii_case("canvas")
    }

    /// Run `f` against the element's 2D context.
    pub fn with_canvas_2d<R>(
        &mut self,
        id: u32,
        f: impl FnOnce(&mut dyn crate::canvas::Canvas) -> R,
    ) -> Option<R> {
        if !self.get_context_2d(id) {
            return None;
        }
        Some(crate::canvas::with_context(self.page, id, f))
    }
}

