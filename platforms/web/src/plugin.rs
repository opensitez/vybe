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
        // The `web:canvas` painter. There is exactly one, so this is an `init`
        // like any other registration — no phase to be careful about. It used to
        // be re-asserted in `finalize` because `platforms/vybe` installed a
        // second painter that resolved through `GuiState` and won on link order;
        // that one is gone (`canvas_backend_impl.rs` deleted), and this install
        // moved back to where it belongs.
        #[cfg(feature = "gui")]
        crate::canvas_backend_widgets::install();
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
                ctor.properties.insert(
                    "name".into(),
                    Value::String(std::sync::Arc::from(name)),
                );
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
        // Binding the handle now (rather than a lazy accessor) is what a
        // browser does too: the document exists before the first script runs.
        // It starts empty, and an empty document opens no window — `should_present`
        // asks `control_count() > 0` — so a console program that never touches
        // `document` is unaffected by its existence.
        //
        // ⛔ It is bound with document id `0` — "the ACTIVE document" — and NOT
        // with `active_document()`. A captured id does not survive: `reset` (and
        // `dom::reset` under it) clears the document map while `next_id` keeps
        // climbing, so a handle taken here names a dead document by the time the
        // program runs, and every call on it goes quiet rather than failing.
        // `doc_arg` resolves 0 to the ambient document at CALL time, which is
        // what `document` means in a browser and what makes one global outlive
        // any number of resets.
        #[cfg(feature = "gui")]
        if let Some(vm) = fw.vm.as_deref_mut() {
            vm.set_global("document", crate::html::document_handle(0));
        }
    }

    /// Drop the widget state the finished program built.
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
    /// ⛔THE WIDGET CRATE IS THE EXCEPTION, and it is why this method exists.
    /// `widgets` keeps its own PROCESS-WIDE tables — `dom::DOCS`,
    /// `ui_events::QUEUE`, `scheduling::TIMERS`/`FRAMES`, each a `static
    /// OnceLock` — and no amount of VM teardown can reach a process global. A
    /// reused VM (the warm pool, `--serve`) would otherwise hand the next
    /// program the previous one's documents, undelivered input, and live
    /// timers.
    ///
    /// This ran in the `vybe` platform while `vybe:gui` owned the widgets. The
    /// GUI is `web:*` now and it is the same `widgets` underneath, so the
    /// obligation moved with it — the plugin that owns the state resets it.
    fn reset(&self) {
        #[cfg(feature = "gui")]
        {
            widgets::dom::reset();
            widgets::ui_events::reset();
            widgets::scheduling::reset();
        }
    }
}

// Link-time registration: this crate submits its plugin to the one registry.
// Nothing lists plugins in code — linking this crate IS the registration.
vybe_runtime::register_plugin!(Plugin);
