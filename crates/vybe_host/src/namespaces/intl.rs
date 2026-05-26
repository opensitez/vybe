//! `Intl.*` namespace — exposes ECMA-402 Intl constructors and static
//! methods as properties of the `Intl` global, matching real JS:
//!
//! ```js
//! new Intl.NumberFormat("en-US").format(1234);
//! Intl.getCanonicalLocales(["EN-us"]);
//! ```
//!
//! Each `Intl.<Class>` resolves to the `new` constructor host fn of the
//! corresponding `ecma:intl/<class>` module. Instance methods (`format`,
//! `compare`, `select`, etc.) dispatch via the TypeRegistry — each Intl
//! class type-stamps its instances with `__type` so the runtime
//! method-resolver finds the right host fn.
//!
//! Static methods (`Intl.getCanonicalLocales`, `Intl.supportedValuesOf`)
//! resolve directly to the `ecma:intl` module.

use super::*;

pub fn register(vm: &mut VM) {
    let intl = ensure_namespace(vm, &["Intl"]);

    // Constructors — `new Intl.X(...)` resolves to `ecma:intl/x:new`.
    let collator = host_fn_ref(vm, "ecma:intl/collator", "new");
    let number_format = host_fn_ref(vm, "ecma:intl/numberformat", "new");
    let date_time_format = host_fn_ref(vm, "ecma:intl/datetimeformat", "new");
    set_prop(&intl, "ListFormat",          host_fn_ref(vm, "ecma:intl/listformat", "new"));
    set_prop(&intl, "PluralRules",         host_fn_ref(vm, "ecma:intl/pluralrules", "new"));
    let relative_time_format = host_fn_ref(vm, "ecma:intl/relativetimeformat", "new");
    let segmenter = host_fn_ref(vm, "ecma:intl/segmenter", "new");
    set_prop(&intl, "Locale",              host_fn_ref(vm, "ecma:intl/locale", "new"));
    set_prop(&intl, "DisplayNames",        host_fn_ref(vm, "ecma:intl/displaynames", "new"));
    set_prop(&intl, "DurationFormat",      host_fn_ref(vm, "ecma:intl/durationformat", "new"));
    set_prop(&intl, "Collator", collator.clone());
    set_prop(&intl, "NumberFormat", number_format.clone());
    set_prop(&intl, "DateTimeFormat", date_time_format.clone());
    set_prop(&intl, "RelativeTimeFormat", relative_time_format.clone());
    set_prop(&intl, "Segmenter", segmenter.clone());

    let object_proto = crate::ecma::object::shared_object_prototype();

    let collator_proto = crate::ecma::intl::shared_collator_prototype();
    set_prop(&collator_proto, "constructor", collator.clone());
    set_prop(&collator_proto, "__proto__", object_proto.clone());
    for name in &["compare", "resolvedOptions"] {
        let idx = *vm.host_registry.get(&("ecma:intl/collator".to_string(), (*name).to_string())).expect("ecma:intl/collator method must be registered");
        set_prop(&collator_proto, name, receiver_host_fn_ref("ecma:intl/collator", name, idx));
    }
    set_prop(&collator, "prototype", collator_proto);

    let number_format_proto = crate::ecma::intl::shared_number_format_prototype();
    set_prop(&number_format_proto, "constructor", number_format.clone());
    set_prop(&number_format_proto, "__proto__", object_proto.clone());
    for name in &["format", "formatToParts", "resolvedOptions"] {
        let idx = *vm.host_registry.get(&("ecma:intl/numberformat".to_string(), (*name).to_string())).expect("ecma:intl/numberformat method must be registered");
        set_prop(&number_format_proto, name, receiver_host_fn_ref("ecma:intl/numberformat", name, idx));
    }
    set_prop(&number_format, "prototype", number_format_proto);

    let date_time_format_proto = crate::ecma::intl::shared_date_time_format_prototype();
    set_prop(&date_time_format_proto, "constructor", date_time_format.clone());
    set_prop(&date_time_format_proto, "__proto__", object_proto.clone());
    for name in &["format", "formatToParts", "resolvedOptions"] {
        let idx = *vm.host_registry.get(&("ecma:intl/datetimeformat".to_string(), (*name).to_string())).expect("ecma:intl/datetimeformat method must be registered");
        set_prop(&date_time_format_proto, name, receiver_host_fn_ref("ecma:intl/datetimeformat", name, idx));
    }
    set_prop(&date_time_format, "prototype", date_time_format_proto);

    let relative_time_format_proto = crate::ecma::intl::shared_relative_time_format_prototype();
    set_prop(&relative_time_format_proto, "constructor", relative_time_format.clone());
    set_prop(&relative_time_format_proto, "__proto__", object_proto.clone());
    for name in &["format", "formatToParts", "resolvedOptions"] {
        let idx = *vm.host_registry.get(&("ecma:intl/relativetimeformat".to_string(), (*name).to_string())).expect("ecma:intl/relativetimeformat method must be registered");
        set_prop(&relative_time_format_proto, name, receiver_host_fn_ref("ecma:intl/relativetimeformat", name, idx));
    }
    set_prop(&relative_time_format, "prototype", relative_time_format_proto);

    let segmenter_proto = crate::ecma::intl::shared_segmenter_prototype();
    set_prop(&segmenter_proto, "constructor", segmenter.clone());
    set_prop(&segmenter_proto, "__proto__", object_proto);
    for name in &["segment", "resolvedOptions"] {
        let idx = *vm.host_registry.get(&("ecma:intl/segmenter".to_string(), (*name).to_string())).expect("ecma:intl/segmenter method must be registered");
        set_prop(&segmenter_proto, name, receiver_host_fn_ref("ecma:intl/segmenter", name, idx));
    }
    set_prop(&segmenter, "prototype", segmenter_proto);

    // Static methods — `Intl.getCanonicalLocales(...)` resolves to
    // `ecma:intl:getCanonicalLocales`.
    set_prop(&intl, "getCanonicalLocales", host_fn_ref(vm, "ecma:intl", "getCanonicalLocales"));
    set_prop(&intl, "supportedValuesOf",   host_fn_ref(vm, "ecma:intl", "supportedValuesOf"));
}
