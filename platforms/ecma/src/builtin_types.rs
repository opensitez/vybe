//! ECMA-262 / ECMA-402 built-in **types** — the runtime TypeRegistry vtables
//! for the JS surface (`String`, `Map`/`Set`/`Weak*`, `RegExp`, `Date`,
//! `Iterator`, `Promise`, and the `Intl.*` classes) plus the universal
//! `Object.prototype` methods on type 0.
//!
//! This is the `register_type` counterpart to the ecma plugin's host-fn
//! `init`: the ecma plugin declares its own types here, in its `finalize`,
//! rather than a central host-side table doing it. Each method resolves a host
//! fn by registry index (via [`Framework::host_fn_index`]), so it must run
//! after every plugin's `init` — which is exactly what `finalize` guarantees.

use vybe_runtime::{Method, TypeDef};
use vybe_runtime::Framework;

/// Register the ECMA/Intl built-in types into the VM's TypeRegistry. Called
/// from the ecma plugin's `finalize`, after `register_globals`.
pub fn register_types(fw: &mut Framework<'_>) {
    // --- Object (type 0, already created) ---
    // Universal methods bound to ECMA-262 primitives. `gethashcode` /
    // `equals` aren't in ECMA — using `String` coercion as a stand-in until a
    // typed equality / hash primitive lands.
    for key in &["tostring", "gethashcode", "equals"] {
        fw.add_host_method(0, key, "ecma:string", "String");
    }
    // §20.1.3 Object.prototype methods — universal on every object so
    // `obj.hasOwnProperty` / `obj.isPrototypeOf` resolve as VALUES (not just
    // as direct calls). The lookup key is lowercased by the TypeRegistry.
    for (key, fname) in &[
        ("hasownproperty", "hasOwnProperty"),
        ("isprototypeof", "isPrototypeOf"),
        ("propertyisenumerable", "propertyIsEnumerable"),
        ("valueof", "valueOf"),
        ("tolocalestring", "toLocaleString"),
    ] {
        fw.add_host_method(0, key, "ecma:object", fname);
    }

    // --- String ---
    //
    // VB/C#/.NET-style instance method names (Contains, IndexOf, Substring,
    // etc.) bound directly to ECMA-262 §22.1 String.prototype. Same import
    // surface JS engines satisfy natively. Methods needing arg adaptation
    // (.NET's `Insert`, `Remove`, `LastIndexOf` with different-shape args)
    // have explicit adapter entries in emitter/dotnet/core/.
    {
        let mut t = TypeDef::new("String");
        for (method, module, fname) in &[
            ("contains", "ecma:string", "includes"),
            ("toupper", "ecma:string", "toUpperCase"),
            ("tolower", "ecma:string", "toLowerCase"),
            // JS-spec camelCase forms — TypeRegistry lowercases the lookup
            // key, so `s.toUpperCase()` and `s.toLowerCase()` both arrive here
            // as `touppercase` / `tolowercase`. These entries make the JS-shape
            // walker rewrites in php/walker.rs (`ucfirst`, `lcfirst`, etc.)
            // resolve through the same ecma:string host fns the JS profile uses.
            ("touppercase", "ecma:string", "toUpperCase"),
            ("tolowercase", "ecma:string", "toLowerCase"),
            ("trim", "ecma:string", "trim"),
            ("trimstart", "ecma:string", "trimStart"),
            ("trimend", "ecma:string", "trimEnd"),
            ("startswith", "ecma:string", "startsWith"),
            ("endswith", "ecma:string", "endsWith"),
            ("indexof", "ecma:string", "indexOf"),
            ("lastindexof", "ecma:string", "lastIndexOf"),
            ("substring", "ecma:string", "substring"),
            ("replace", "ecma:string", "replace"),
            ("split", "ecma:string", "split"),
            ("padleft", "ecma:string", "padStart"),
            ("padright", "ecma:string", "padEnd"),
            ("tostring", "ecma:string", "String"),
            ("toupperinvariant", "ecma:string", "toUpperCase"),
            ("tolowerinvariant", "ecma:string", "toLowerCase"),
            ("chars", "ecma:string", "charAt"),
            ("insert", "ecma:string", "substring"),
            ("remove", "ecma:string", "slice"),
            // ── JS-spec methods exposed under the same ecma:string surface so
            // `s.charAt(...)`, `s.charCodeAt(...)`, etc. dispatch via the
            // TypeRegistry without language-specific aliasing.
            ("charat", "ecma:string", "charAt"),
            ("charcodeat", "ecma:string", "charCodeAt"),
            ("codePointAt", "ecma:string", "codePointAt"),
            ("at", "ecma:string", "at"),
            ("concat", "ecma:string", "concat"),
            ("includes", "ecma:string", "includes"),
            ("repeat", "ecma:string", "repeat"),
            ("padstart", "ecma:string", "padStart"),
            ("padend", "ecma:string", "padEnd"),
            ("normalize", "ecma:string", "normalize"),
            ("slice", "ecma:string", "slice"),
            ("substr", "ecma:string", "substr"),
            ("trimleft", "ecma:string", "trimStart"),
            ("trimright", "ecma:string", "trimEnd"),
            ("localecompare", "ecma:string", "localeCompare"),
            ("tolocalelowercase", "ecma:string", "toLowerCase"),
            ("tolocaleuppercase", "ecma:string", "toUpperCase"),
            // RegExp-driven String.prototype methods (§22.1.3.{11,13,14,17,18}).
            ("match", "ecma:regexp", "match"),
            ("matchAll", "ecma:regexp", "matchAll"),
            ("search", "ecma:regexp", "search"),
            ("replaceall", "ecma:regexp", "replaceAll"),
            // §22.1.5.1: String.prototype[@@iterator]() — yields code points.
            // Strings are opaque in WASM (wasm:js-string spec) so iteration
            // requires a host function, same as charCodeAt/codePointAt.
            ("Symbol(@@iterator)", "ecma:string", "iterator"),
        ] {
            if let Some(idx) = fw.host_fn_index(module, fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0); // inherits from Object
        fw.register_type(t);
    }

    // --- Map (JS collection) ---
    //
    // Bound directly to `ecma:map.*` — same imports JS engines satisfy
    // natively. Backing store is `ObjectKind::Map(IndexMap<Value,Value>)`
    // (SameValueZero keys per ECMA-262 §24.1).
    {
        let mut t = TypeDef::new("Map");
        for (method, fname) in &[
            ("set", "set"),
            ("get", "get"),
            ("has", "has"),
            ("delete", "delete"),
            ("keys", "keys"),
            ("values", "values"),
            ("clear", "clear"),
            ("entries", "entries"),
            ("forEach", "forEach"),
            // §24.1.3.12: Map.prototype[@@iterator] = Map.prototype.entries
            ("Symbol(@@iterator)", "entries"),
        ] {
            if let Some(idx) = fw.host_fn_index("ecma:map", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        // `size` is a property updated by every mutator, but expose the method
        // form too for VB-style `.Count` callers and JS code that does
        // `m.size()` even though spec says it's a getter.
        if let Some(idx) = fw.host_fn_index("ecma:map", "size") {
            t.methods.insert("size".to_string(), Method::HostFn(idx));
            t.methods.insert("count".to_string(), Method::HostFn(idx));
        }
        t.parent = Some(0);
        fw.register_type(t);
    }

    // --- Set (JS collection) ---
    //
    // Bound to `ecma:set.*`. Backing store is `ObjectKind::Set(IndexSet<Value>)`.
    {
        let mut t = TypeDef::new("Set");
        for (method, fname) in &[
            ("add", "add"),
            ("has", "has"),
            ("delete", "delete"),
            ("values", "values"),
            ("keys", "values"),
            ("clear", "clear"),
            ("entries", "entries"),
            ("forEach", "forEach"),
            // §24.2.3.11: Set.prototype[@@iterator] = Set.prototype.values
            ("Symbol(@@iterator)", "values"),
        ] {
            if let Some(idx) = fw.host_fn_index("ecma:set", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        if let Some(idx) = fw.host_fn_index("ecma:set", "size") {
            t.methods.insert("size".to_string(), Method::HostFn(idx));
            t.methods.insert("count".to_string(), Method::HostFn(idx));
        }
        t.parent = Some(0);
        fw.register_type(t);
    }

    // --- WeakMap (JS collection) ---
    //
    // Bound to `ecma:weakmap.*` — share-nothing with Map (different backing
    // shape: parallel keys/values arrays per ES spec GC model).
    {
        let mut t = TypeDef::new("WeakMap");
        for (method, fname) in &[
            ("set", "set"),
            ("get", "get"),
            ("has", "has"),
            ("delete", "delete"),
        ] {
            if let Some(idx) = fw.host_fn_index("ecma:weakmap", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        fw.register_type(t);
    }

    // --- WeakSet (JS collection) ---
    {
        let mut t = TypeDef::new("WeakSet");
        for (method, fname) in &[("add", "add"), ("has", "has"), ("delete", "delete")] {
            if let Some(idx) = fw.host_fn_index("ecma:weakset", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        fw.register_type(t);
    }

    // --- WeakRef (ECMA-262 §26.1) ---
    //
    // Bound to `ecma:weakref.*`. Strong-ref MVP stand-in (WASM GC MVP doesn't
    // expose weak refs); `deref()` always returns the target.
    {
        let mut t = TypeDef::new("WeakRef");
        for (method, fname) in &[("deref", "deref")] {
            if let Some(idx) = fw.host_fn_index("ecma:weakref", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        fw.register_type(t);
    }

    // --- FinalizationRegistry (ECMA-262 §26.2) ---
    //
    // Bound to `ecma:finalizationregistry.*`. Strong-ref backing means the
    // cleanup callback never fires — API surface only.
    {
        let mut t = TypeDef::new("FinalizationRegistry");
        for (method, fname) in &[("register", "register"), ("unregister", "unregister")] {
            if let Some(idx) = fw.host_fn_index("ecma:finalizationregistry", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        fw.register_type(t);
    }

    // --- RegExp (ECMA-262 §22.2) ---
    //
    // Bound to `ecma:regexp.*`. Backing object is a plain Object with
    // `source`/`flags`/etc. as own properties — re-compiled at every call.
    {
        let mut t = TypeDef::new("RegExp");
        for (method, fname) in &[("test", "test"), ("exec", "exec"), ("toString", "toString")] {
            if let Some(idx) = fw.host_fn_index("ecma:regexp", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        fw.register_type(t);
    }

    // --- Date (ECMA-262 §21.4) ---
    //
    // Backing object is a plain Object with `__time` (ms since epoch). Methods
    // all read/write that property via `ecma:date.*`. The TypeDef entry makes
    // `d.getFullYear()` resolve through the vtable instead of falling through
    // to instance-property lookup. `resolve_method` lowercases lookups.
    {
        let mut t = TypeDef::new("Date");
        for method in &[
            "getFullYear",
            "getYear",
            "getMonth",
            "getDate",
            "getDay",
            "getHours",
            "getMinutes",
            "getSeconds",
            "getMilliseconds",
            "getUTCFullYear",
            "getUTCMonth",
            "getUTCDate",
            "getUTCDay",
            "getUTCHours",
            "getUTCMinutes",
            "getUTCSeconds",
            "getUTCMilliseconds",
            "getTime",
            "getTimezoneOffset",
            "valueOf",
            "setTime",
            "setFullYear",
            "setMonth",
            "setDate",
            "setHours",
            "setMinutes",
            "setSeconds",
            "setMilliseconds",
            "setUTCFullYear",
            "setUTCMonth",
            "setUTCDate",
            "setUTCHours",
            "setUTCMinutes",
            "setUTCSeconds",
            "setUTCMilliseconds",
            "toISOString",
            "toString",
            "toUTCString",
            "toDateString",
            "toTimeString",
            "toJSON",
        ] {
            if let Some(idx) = fw.host_fn_index("ecma:date", method) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        fw.register_type(t);
    }

    // --- Intl.* (ECMA-402) ---
    //
    // Each Intl class is bound to its `ecma:intl/<class>` host module.
    {
        let mut t = TypeDef::new("Collator");
        for (method, fname) in &[
            ("compare", "compare"),
            ("resolvedOptions", "resolvedOptions"),
        ] {
            if let Some(idx) = fw.host_fn_index("ecma:intl/collator", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        fw.register_type(t);
    }
    {
        let mut t = TypeDef::new("NumberFormat");
        for (method, fname) in &[
            ("format", "format"),
            ("formatToParts", "formatToParts"),
            ("resolvedOptions", "resolvedOptions"),
        ] {
            if let Some(idx) = fw.host_fn_index("ecma:intl/numberformat", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        fw.register_type(t);
    }
    {
        let mut t = TypeDef::new("DateTimeFormat");
        for (method, fname) in &[
            ("format", "format"),
            ("formatToParts", "formatToParts"),
            ("resolvedOptions", "resolvedOptions"),
        ] {
            if let Some(idx) = fw.host_fn_index("ecma:intl/datetimeformat", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        fw.register_type(t);
    }
    {
        let mut t = TypeDef::new("ListFormat");
        for (method, fname) in &[
            ("format", "format"),
            ("formatToParts", "formatToParts"),
            ("resolvedOptions", "resolvedOptions"),
        ] {
            if let Some(idx) = fw.host_fn_index("ecma:intl/listformat", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        fw.register_type(t);
    }
    {
        let mut t = TypeDef::new("PluralRules");
        for (method, fname) in &[
            ("select", "select"),
            ("selectRange", "selectRange"),
            ("resolvedOptions", "resolvedOptions"),
        ] {
            if let Some(idx) = fw.host_fn_index("ecma:intl/pluralrules", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        fw.register_type(t);
    }
    {
        let mut t = TypeDef::new("RelativeTimeFormat");
        for (method, fname) in &[
            ("format", "format"),
            ("formatToParts", "formatToParts"),
            ("resolvedOptions", "resolvedOptions"),
        ] {
            if let Some(idx) = fw.host_fn_index("ecma:intl/relativetimeformat", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        fw.register_type(t);
    }
    {
        let mut t = TypeDef::new("Segmenter");
        for (method, fname) in &[
            ("segment", "segment"),
            ("resolvedOptions", "resolvedOptions"),
        ] {
            if let Some(idx) = fw.host_fn_index("ecma:intl/segmenter", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        fw.register_type(t);
    }
    {
        let mut t = TypeDef::new("Locale");
        for (method, fname) in &[
            ("toString", "toString"),
            ("maximize", "maximize"),
            ("minimize", "minimize"),
        ] {
            if let Some(idx) = fw.host_fn_index("ecma:intl/locale", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        fw.register_type(t);
    }
    {
        let mut t = TypeDef::new("DisplayNames");
        for (method, fname) in &[("of", "of"), ("resolvedOptions", "resolvedOptions")] {
            if let Some(idx) = fw.host_fn_index("ecma:intl/displaynames", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        fw.register_type(t);
    }
    {
        let mut t = TypeDef::new("DurationFormat");
        for (method, fname) in &[
            ("format", "format"),
            ("formatToParts", "formatToParts"),
            ("resolvedOptions", "resolvedOptions"),
        ] {
            if let Some(idx) = fw.host_fn_index("ecma:intl/durationformat", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        fw.register_type(t);
    }

    // --- Promise (JS) ---
    {
        let t = TypeDef::new("Promise");
        fw.register_type(t);
    }

    // --- Iterator (Stage-3 iterator helpers) ---
    {
        let mut t = TypeDef::new("Iterator");
        for (method, fname) in &[
            ("take", "take"),
            ("drop", "drop"),
            ("map", "map"),
            ("filter", "filter"),
            ("reduce", "reduce"),
            ("forEach", "forEach"),
            ("some", "some"),
            ("every", "every"),
            ("find", "find"),
            ("toArray", "toArray"),
            ("flatMap", "flatMap"),
        ] {
            if let Some(idx) = fw.host_fn_index("ecma:iterator", fname) {
                t.methods.insert(method.to_string(), Method::HostFn(idx));
            }
        }
        t.parent = Some(0);
        fw.register_type(t);
    }
}
