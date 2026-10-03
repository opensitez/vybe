//! The web platform as a `vybe_runtime::Plugin` — one plugin, same type as all
//! the others. `init` registers the `web:*` host functions (WHATWG / W3C:
//! crypto, URL, TextEncoder, fetch, dom-parser). Always-on (pure computation).

use vybe_runtime::Value;

/// The web platform plugin.
pub struct Plugin;

impl vybe_runtime::Plugin for Plugin {
    fn name(&self) -> &'static str {
        "web"
    }

    fn init(&self, fw: &mut vybe_runtime::Framework<'_>) {
        if let Some(vm) = fw.vm.as_deref_mut() {
            crate::register(vm);
        }
        // `crate::register` installs the selected browser engine and its
        // matching canvas backend together. Re-installing a concrete backend
        // here could split DOM operations and canvas draws across engines.
    }

    fn finalize(&self, fw: &mut vybe_runtime::Framework<'_>) {
        // The web-platform TypeRegistry vtables (TextEncoder/Decoder,
        // URLSearchParams, Response, DOM node hierarchy) + DOM type-id
        // stamping — registered via the `register_type` primitive after every
        // plugin's host fns exist.
        crate::builtin_types::register_types(fw);

        // ── WHATWG constructor ↔ prototype wiring ──────────────────────
        //
        // `URL`, `URLSearchParams`, `TextEncoder` and `TextDecoder` are
        // CONSTRUCTORS, and a constructor is a function with a `prototype` its
        // instances link to. They had neither: the bare name resolved to a
        // namespace placeholder (`typeof URL === "object"`), so
        // `u instanceof URL` had nothing to walk and was answered by a `__type`
        // string compare inside the ECMA host — a vybe stamp deciding identity,
        // which is exactly what the ECMA host is not allowed to do.
        //
        // `__ctor_<Name>` is the anchor the compiler resolves a bare read
        // through. ⛔ A name published here MUST also appear in the compiler's
        // `is_js_builtin_ctor_value`, or the bare read resolves to the
        // placeholder again.
        if let Some(vm) = fw.vm.as_deref_mut() {
            for (name, module, ctor_fn, proto) in [
                ("URL", "web:url", "new", crate::url::shared_url_prototype()),
                (
                    "URLSearchParams",
                    "web:url",
                    "searchParamsNew",
                    crate::url::shared_url_search_params_prototype(),
                ),
                (
                    "TextEncoder",
                    "web:encoding",
                    "encoderNew",
                    crate::encoding::shared_text_encoder_prototype(),
                ),
                (
                    "TextDecoder",
                    "web:encoding",
                    "decoderNew",
                    crate::encoding::shared_text_decoder_prototype(),
                ),
            ] {
                let Some(&idx) = vm
                    .host_registry
                    .get(&(module.to_string(), ctor_fn.to_string()))
                else {
                    continue;
                };
                let mut ctor = vybe_runtime::value::Object::new();
                ctor.kind = vybe_runtime::value::ObjectKind::HostFunction(idx);
                ctor.properties
                    .insert("name".into(), Value::String(std::sync::Arc::from(name)));
                ctor.properties.insert("prototype".into(), proto.clone());
                let ctor = Value::Object(vybe_runtime::heap::alloc(ctor));
                if let Value::Object(p) = &proto {
                    p.lock()
                        .unwrap()
                        .properties
                        .insert("constructor".into(), ctor.clone());
                }
                vm.set_global_owned(name.to_string(), ctor.clone());
                vm.set_global_owned(format!("__ctor_{name}"), ctor);
            }
        }

        // `document` — a property of the global object, which is where a
        // browser puts it (HTML §7.3: `window.document`). Guest code says
        // `document.createElement("button")` and means the document it is
        // running in; there is nothing to import and nothing to construct.
        //
        // It is created HERE, not in `init`, because the handle carries the
        // `HTMLDocument` type id that `register_types` above has only just
        // assigned. Stamping it a phase earlier would leave it type 0 and every
        // method call on it unresolvable.
        //
        // Bind the handle now, but resolve its ambient document and body only
        // when guest code calls into the DOM. Console programs stay headless.
        //
        // ⛔ It is bound with document id `0` — "the ACTIVE document" — and NOT
        // with `active_document()`. A captured id does not survive: `reset` (and
        // `dom::reset` under it) clears the document map while `next_id` keeps
        // climbing, so a handle taken here names a dead document by the time the
        // program runs, and every call on it goes quiet rather than failing.
        // `doc_arg` resolves 0 to the ambient document at CALL time, which is
        // what `document` means in a browser and what makes one global outlive
        // any number of resets.
        #[cfg(feature = "engine-webcore")]
        if let Some(vm) = fw.vm.as_deref_mut() {
            let body_getter = vm
                .resolve_host_function_index("web:html", "body")
                .expect("web:html.body is registered");
            vm.set_global("document", crate::html::document_handle(0, body_getter));
        }
    }

    /// Drop browser state the finished program built.
    ///
    /// Most of what this platform holds needs nothing here: the DOM listener
    /// table and the ambient document are VM-owned storage
    /// (`vybe_runtime::resources`), so `reset_to` drops them without this
    /// plugin taking part. That is the whole point of the store — the listener
    /// table used to be a process-global static with a hand-written
    /// `reset_listeners`, and `reset_active_document`'s only caller was a
    /// pascal test helper; a per-test helper cannot be the mechanism, it fixes
    /// the one caller that remembers to call it and leaves every other embedder
    /// broken. Queued timer and animation callbacks are not here either: they
    /// are `DeferredSource`s, and `reset_to` clears every registered source's
    /// queue through `clear_pending`.
    ///
    fn reset(&self) {
        #[cfg(feature = "engine-webcore")]
        if crate::engine_select::live() == Some(crate::engine_select::Engine::WebCore) {
            webcore::ui_events::reset();
            webcore::scheduling::reset();
            return;
        }
        #[cfg(feature = "engine-osbrowser")]
        if crate::engine_select::live() == Some(crate::engine_select::Engine::OsBrowser) {
            crate::engine_osbrowser::reset();
            return;
        }
    }
}

// Link-time registration: this crate submits its plugin to the one registry.
// Nothing lists plugins in code — linking this crate IS the registration.
vybe_runtime::register_plugin!(Plugin);
