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
    set_prop(&intl, "Collator",            host_fn_ref(vm, "ecma:intl/collator", "new"));
    set_prop(&intl, "NumberFormat",        host_fn_ref(vm, "ecma:intl/numberformat", "new"));
    set_prop(&intl, "DateTimeFormat",      host_fn_ref(vm, "ecma:intl/datetimeformat", "new"));
    set_prop(&intl, "ListFormat",          host_fn_ref(vm, "ecma:intl/listformat", "new"));
    set_prop(&intl, "PluralRules",         host_fn_ref(vm, "ecma:intl/pluralrules", "new"));
    set_prop(&intl, "RelativeTimeFormat",  host_fn_ref(vm, "ecma:intl/relativetimeformat", "new"));
    set_prop(&intl, "Segmenter",           host_fn_ref(vm, "ecma:intl/segmenter", "new"));
    set_prop(&intl, "Locale",              host_fn_ref(vm, "ecma:intl/locale", "new"));
    set_prop(&intl, "DisplayNames",        host_fn_ref(vm, "ecma:intl/displaynames", "new"));
    set_prop(&intl, "DurationFormat",      host_fn_ref(vm, "ecma:intl/durationformat", "new"));

    // Static methods — `Intl.getCanonicalLocales(...)` resolves to
    // `ecma:intl:getCanonicalLocales`.
    set_prop(&intl, "getCanonicalLocales", host_fn_ref(vm, "ecma:intl", "getCanonicalLocales"));
    set_prop(&intl, "supportedValuesOf",   host_fn_ref(vm, "ecma:intl", "supportedValuesOf"));
}
