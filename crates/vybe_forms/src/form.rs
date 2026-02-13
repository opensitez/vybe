use crate::control::Control;
use crate::events::EventBinding;
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Form {
    pub name: String,
    pub text: String,
    pub width: i32,
    pub height: i32,
    pub controls: Vec<Control>,
    pub event_bindings: Vec<EventBinding>,
    #[serde(default)]
    pub back_color: Option<String>,
    #[serde(default)]
    pub fore_color: Option<String>,
    #[serde(default)]
    pub font: Option<String>,
}

impl Form {
    pub fn new(name: impl Into<String>) -> Self {
        let name = name.into();
        Self {
            name: name.clone(),
            text: name,
            width: 640,
            height: 480,
            controls: Vec::new(),
            event_bindings: Vec::new(),
            back_color: None,
            fore_color: None,
            font: None,
        }
    }

    pub fn add_control(&mut self, control: Control) {
        self.controls.push(control);
    }

    pub fn remove_control(&mut self, id: uuid::Uuid) {
        self.controls.retain(|c| c.id != id);
    }

    pub fn get_control(&self, id: uuid::Uuid) -> Option<&Control> {
        self.controls.iter().find(|c| c.id == id)
    }

    pub fn get_control_mut(&mut self, id: uuid::Uuid) -> Option<&mut Control> {
        self.controls.iter_mut().find(|c| c.id == id)
    }

    pub fn get_control_by_name(&self, name: &str) -> Option<&Control> {
        self.controls.iter().find(|c| c.name.eq_ignore_ascii_case(name))
    }

    pub fn get_control_by_name_mut(&mut self, name: &str) -> Option<&mut Control> {
        self.controls.iter_mut().find(|c| c.name.eq_ignore_ascii_case(name))
    }

    pub fn find_control_at(&self, x: i32, y: i32) -> Option<&Control> {
        // Find the topmost control at the given position
        // Controls added later are on top
        self.controls.iter().rev().find(|c| c.bounds.contains(x, y))
    }

    pub fn add_event_binding(&mut self, binding: EventBinding) {
        self.event_bindings.push(binding);
    }

    pub fn get_event_handler(&self, control_name: &str, event_type: &crate::events::EventType) -> Option<&str> {
        self.event_bindings
            .iter()
            .find(|b| b.control_name == control_name && &b.event_type == event_type)
            .map(|b| b.handler_name.as_str())
    }

    pub fn get_control_array(&self, name: &str) -> Vec<&Control> {
        self.controls.iter()
            .filter(|c| c.name.eq_ignore_ascii_case(name) && c.index.is_some())
            .collect()
    }

    pub fn is_control_array(&self, name: &str) -> bool {
        self.controls.iter().any(|c| c.name.eq_ignore_ascii_case(name) && c.index.is_some())
    }

    pub fn next_array_index(&self, name: &str) -> i32 {
        self.controls.iter()
            .filter(|c| c.name.eq_ignore_ascii_case(name) && c.index.is_some())
            .filter_map(|c| c.index)
            .max()
            .map(|m| m + 1)
            .unwrap_or(0)
    }
}
