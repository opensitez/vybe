//! `ecma:intl/*` — ECMA-402 Internationalization API.
//!
//! Reference: <https://tc39.es/ecma402/>.
//!
//! Each Intl class becomes its own host module since Vybe's host
//! registry is keyed by `(module, name)`:
//!
//!   - `ecma:intl/collator`           — Intl.Collator           (icu::collator)
//!   - `ecma:intl/numberformat`       — Intl.NumberFormat       (icu::decimal)
//!   - `ecma:intl/datetimeformat`     — Intl.DateTimeFormat     (icu::datetime)
//!   - `ecma:intl/listformat`         — Intl.ListFormat         (icu::list)
//!   - `ecma:intl/pluralrules`        — Intl.PluralRules        (intl_pluralrules — focused crate, full CLDR)
//!   - `ecma:intl/relativetimeformat` — Intl.RelativeTimeFormat (icu_relativetime)
//!   - `ecma:intl/segmenter`          — Intl.Segmenter          (unicode-segmentation — focused crate, UAX #29)
//!   - `ecma:intl/locale`             — Intl.Locale             (unic-langid + icu::locale)
//!   - `ecma:intl/displaynames`       — Intl.DisplayNames       (icu_displaynames)
//!   - `ecma:intl/durationformat`     — Intl.DurationFormat     (MVP — icu_experimental still pre-1.0)
//!   - `ecma:intl`                    — static fns (getCanonicalLocales, supportedValuesOf)
//!
//! Hybrid backing — focused crates (`intl_pluralrules`, `unicode-segmentation`,
//! `unic-langid`) for services that have full multi-locale coverage at
//! small size; ICU4X 2.x umbrella `icu` crate for the rest. Same pattern
//! Node small-icu uses (real Intl code, locale-aware behaviour).

use std::str::FromStr;
use std::sync::{Arc, Mutex, OnceLock};
use vybe_runtime::value::{Object, ObjectKind, Value};
use vybe_runtime::{HostContext, VM};

static COLLATOR_PROTOTYPE: OnceLock<Arc<Mutex<Object>>> = OnceLock::new();
static NUMBER_FORMAT_PROTOTYPE: OnceLock<Arc<Mutex<Object>>> = OnceLock::new();
static DATE_TIME_FORMAT_PROTOTYPE: OnceLock<Arc<Mutex<Object>>> = OnceLock::new();
static RELATIVE_TIME_FORMAT_PROTOTYPE: OnceLock<Arc<Mutex<Object>>> = OnceLock::new();
static SEGMENTER_PROTOTYPE: OnceLock<Arc<Mutex<Object>>> = OnceLock::new();

fn bound_host_fn_ref_by_idx(module: &str, name: &str, idx: usize, bound_args: Vec<Value>) -> Value {
    let mut obj = Object::new();
    obj.properties
        .insert("__host_module".into(), Value::String(Arc::from(module)));
    obj.properties
        .insert("__host_name".into(), Value::String(Arc::from(name)));
    obj.properties
        .insert("__host_idx".into(), Value::F64(idx as f64));
    obj.properties
        .insert("name".into(), Value::String(Arc::from(name)));
    obj.properties.insert(
        "__bound_args".into(),
        Value::Object(vybe_runtime::heap::alloc(Object::new_array(bound_args))),
    );
    obj.kind = ObjectKind::HostFunction(idx);
    Value::Object(vybe_runtime::heap::alloc(obj))
}

pub fn register(vm: &mut VM) {
    register_collator(vm);
    register_number_format(vm);
    register_date_time_format(vm);
    register_list_format(vm);
    register_plural_rules(vm);
    register_relative_time_format(vm);
    register_segmenter(vm);
    register_locale(vm);
    register_display_names(vm);
    register_duration_format(vm);
    register_static(vm);
}

// ── Common helpers ───────────────────────────────────────────────────

fn s_val(text: &str) -> Value {
    Value::String(Arc::from(text))
}

pub fn shared_collator_prototype() -> Value {
    Value::Object(
        COLLATOR_PROTOTYPE
            .get_or_init(|| vybe_runtime::heap::alloc(Object::new()))
            .clone(),
    )
}

pub fn shared_number_format_prototype() -> Value {
    Value::Object(
        NUMBER_FORMAT_PROTOTYPE
            .get_or_init(|| vybe_runtime::heap::alloc(Object::new()))
            .clone(),
    )
}

pub fn shared_date_time_format_prototype() -> Value {
    Value::Object(
        DATE_TIME_FORMAT_PROTOTYPE
            .get_or_init(|| vybe_runtime::heap::alloc(Object::new()))
            .clone(),
    )
}

pub fn shared_relative_time_format_prototype() -> Value {
    Value::Object(
        RELATIVE_TIME_FORMAT_PROTOTYPE
            .get_or_init(|| vybe_runtime::heap::alloc(Object::new()))
            .clone(),
    )
}

pub fn shared_segmenter_prototype() -> Value {
    Value::Object(
        SEGMENTER_PROTOTYPE
            .get_or_init(|| vybe_runtime::heap::alloc(Object::new()))
            .clone(),
    )
}

/// Unicode locale extension keywords from a BCP-47 tag's `-u-` sequence
/// (UTS #35): `en-US-u-ca-buddhist-hc-h12` → `{ca: "buddhist", hc: "h12"}`.
///
/// A two-letter subtag starts a new key; everything up to the next key is that
/// key's (possibly multi-subtag) type. A key with no type is the boolean form
/// (`-u-kn-` means `kn=true`), which is why an empty value is meaningful.
fn unicode_extension_keywords(tag: &str) -> std::collections::HashMap<String, String> {
    let mut out = std::collections::HashMap::new();
    let lower = tag.to_ascii_lowercase();
    let mut parts = lower.split('-');
    // Advance to the `u` singleton; anything before it is the language id.
    if !parts.any(|p| p == "u") {
        return out;
    }
    let mut key: Option<String> = None;
    let mut value: Vec<String> = Vec::new();
    for part in parts {
        // Another singleton ends the unicode extension entirely.
        if part.len() == 1 {
            break;
        }
        if part.len() == 2 {
            if let Some(k) = key.take() {
                out.insert(k, value.join("-"));
                value.clear();
            }
            key = Some(part.to_string());
        } else if key.is_some() {
            value.push(part.to_string());
        }
    }
    if let Some(k) = key {
        out.insert(k, value.join("-"));
    }
    out
}

fn make_array(elements: Vec<Value>) -> Value {
    let mut obj = Object::new_array(elements);
    obj.properties
        .insert("__type".into(), Value::String(Arc::from("Array")));
    obj.properties
        .insert("__proto__".into(), crate::array::shared_array_prototype());
    Value::Object(vybe_runtime::heap::alloc(obj))
}

fn make_object(props: Vec<(&str, Value)>) -> Value {
    let mut obj = Object::new();
    for (k, v) in props {
        obj.properties.insert(k.into(), v);
    }
    Value::Object(vybe_runtime::heap::alloc(obj))
}

fn part_obj(part_type: &str, value: &str) -> Value {
    make_object(vec![("type", s_val(part_type)), ("value", s_val(value))])
}

fn obj_string_prop(obj: &Arc<Mutex<Object>>, key: &str) -> Option<String> {
    let lock = obj.lock().unwrap();
    match lock.properties.get(key)? {
        Value::String(s) => Some(s.to_string()),
        other => Some(format!("{}", other)),
    }
}

/// Resolve the locale arg from `args[0]` — accepts a string tag or an
/// array of tags (per ECMA-402, the first supported tag wins). Default
/// when missing or empty: "en-US".
fn resolve_locale(arg: Option<&Value>) -> String {
    canonicalize_locale(&match arg {
        Some(Value::String(tag)) if !tag.is_empty() => tag.to_string(),
        Some(Value::Object(o)) => {
            let lock = o.lock().unwrap();
            if let ObjectKind::Array(ref elems) = lock.kind {
                if let Some(Value::String(tag)) = elems.first() {
                    return canonicalize_locale(tag);
                }
            }
            "en-US".into()
        }
        _ => "en-US".into(),
    })
}

fn canonicalize_locale(tag: &str) -> String {
    unic_langid::LanguageIdentifier::from_str(tag)
        .map(|id| id.to_string())
        .unwrap_or_else(|_| tag.to_string())
}

fn is_invalid_locale_tag(tag: &str) -> bool {
    tag.is_empty() || unic_langid::LanguageIdentifier::from_str(tag).is_err()
}

fn option_string(obj: &Object, key: &str) -> Option<String> {
    obj.properties.get(key).and_then(|v| match v {
        Value::String(s) => Some(s.to_string()),
        _ => None,
    })
}

fn resolve_options(arg: Option<&Value>) -> Arc<Mutex<Object>> {
    if let Some(Value::Object(o)) = arg {
        o.clone()
    } else {
        vybe_runtime::heap::alloc(Object::new())
    }
}

fn resolve_date_ms(arg: Option<&Value>) -> Option<f64> {
    match arg {
        Some(Value::Object(obj)) => {
            let lock = obj.lock().unwrap();
            lock.properties.get("__time").map(|v| v.as_f64())
        }
        Some(value) => Some(value.as_f64()),
        None => None,
    }
}

/// Parse a tag into a `unic_langid::LanguageIdentifier`. Falls back to
/// `en-US` on parse failure (matches ECMA-402's "best fit" behaviour).
fn parse_langid(tag: &str) -> unic_langid::LanguageIdentifier {
    unic_langid::LanguageIdentifier::from_str(tag)
        .unwrap_or_else(|_| unic_langid::LanguageIdentifier::from_str("en-US").unwrap())
}

/// Same but produces `icu::locale::Locale` for the ICU4X-backed services.
fn parse_icu_locale(tag: &str) -> icu::locale::Locale {
    icu::locale::Locale::from_str(tag)
        .unwrap_or_else(|_| icu::locale::Locale::from_str("en-US").unwrap())
}

// ── Intl.Collator (ECMA-402 §11) ────────────────────────────────────

/// Digit-run-aware comparison for `Intl.Collator(…, { numeric: true })`:
/// consecutive ASCII digits compare as numbers, everything else as chars.
fn natural_compare(a: &str, b: &str) -> std::cmp::Ordering {
    let (ab, bb) = (a.as_bytes(), b.as_bytes());
    let (mut i, mut j) = (0usize, 0usize);
    while i < ab.len() && j < bb.len() {
        if ab[i].is_ascii_digit() && bb[j].is_ascii_digit() {
            let i0 = i;
            let j0 = j;
            while i < ab.len() && ab[i].is_ascii_digit() {
                i += 1;
            }
            while j < bb.len() && bb[j].is_ascii_digit() {
                j += 1;
            }
            let na = a[i0..i].trim_start_matches('0');
            let nb = b[j0..j].trim_start_matches('0');
            let ord = na.len().cmp(&nb.len()).then_with(|| na.cmp(nb));
            if ord != std::cmp::Ordering::Equal {
                return ord;
            }
        } else {
            let ord = ab[i].cmp(&bb[j]);
            if ord != std::cmp::Ordering::Equal {
                return ord;
            }
            i += 1;
            j += 1;
        }
    }
    (ab.len() - i).cmp(&(bb.len() - j))
}

fn register_collator(vm: &mut VM) {
    // Register compare first so we can look up its index for bound instances.
    vm.register_host_fn(
        "ecma:intl/collator",
        "compare",
        Box::new(|ctx, args| {
            use icu::collator::{Collator, options::CollatorOptions};
            let collator = match args.first() {
                Some(Value::Object(o)) => o.clone(),
                _ => return Value::I32(0),
            };
            if matches!(args.get(1), Some(Value::Symbol(_)))
                || matches!(args.get(2), Some(Value::Symbol(_)))
            {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "TypeError",
                    "Cannot convert Symbol to string",
                ));
                return Value::Undefined;
            }
            let mut a = match args.get(1) {
                Some(Value::String(s)) => s.to_string(),
                Some(o) => format!("{}", o),
                None => String::new(),
            };
            let mut b = match args.get(2) {
                Some(Value::String(s)) => s.to_string(),
                Some(o) => format!("{}", o),
                None => String::new(),
            };
            let locale = obj_string_prop(&collator, "locale").unwrap_or_else(|| "en-US".into());
            let sensitivity =
                obj_string_prop(&collator, "sensitivity").unwrap_or_else(|| "variant".into());
            let ignore_punctuation = matches!(
                collator.lock().unwrap().properties.get("ignorePunctuation"),
                Some(Value::Bool(true))
            );
            if ignore_punctuation {
                a.retain(|c| !c.is_ascii_punctuation());
                b.retain(|c| !c.is_ascii_punctuation());
            }
            let case_first =
                obj_string_prop(&collator, "caseFirst").unwrap_or_else(|| "false".into());
            if case_first == "upper" && a.eq_ignore_ascii_case(&b) && a != b {
                let a_upper = a.chars().next().map(|c| c.is_uppercase()).unwrap_or(false);
                let b_upper = b.chars().next().map(|c| c.is_uppercase()).unwrap_or(false);
                if a_upper != b_upper {
                    return Value::I32(if a_upper { -1 } else { 1 });
                }
            }
            let strip_accents = |s: &str| {
                use unicode_normalization::UnicodeNormalization;
                s.nfd()
                    .filter(|c| !('\u{0300}'..='\u{036f}').contains(c))
                    .collect::<String>()
            };
            match sensitivity.as_str() {
                "base" => {
                    if strip_accents(&a).to_lowercase() == strip_accents(&b).to_lowercase() {
                        return Value::I32(0);
                    }
                }
                "accent" => {
                    if a.to_lowercase() == b.to_lowercase() {
                        return Value::I32(0);
                    }
                }
                _ => {}
            }
            // ECMA-402 §10.1.2 numeric collation: digit runs compare by
            // numeric value ("file2" < "file10").
            let numeric = matches!(
                collator.lock().unwrap().properties.get("numeric"),
                Some(Value::Bool(true))
            );
            if numeric
                && a.bytes().any(|b| b.is_ascii_digit())
                && b.bytes().any(|b| b.is_ascii_digit())
            {
                return Value::I32(match natural_compare(&a, &b) {
                    std::cmp::Ordering::Less => -1,
                    std::cmp::Ordering::Equal => 0,
                    std::cmp::Ordering::Greater => 1,
                });
            }
            let icu_loc = parse_icu_locale(&locale);
            let prefs = (&icu_loc).into();
            let coll = match Collator::try_new(prefs, CollatorOptions::default()) {
                Ok(c) => c,
                Err(_) => {
                    return Value::I32(match a.as_str().cmp(b.as_str()) {
                        std::cmp::Ordering::Less => -1,
                        std::cmp::Ordering::Equal => 0,
                        std::cmp::Ordering::Greater => 1,
                    });
                }
            };
            Value::I32(match coll.compare(&a, &b) {
                std::cmp::Ordering::Less => -1,
                std::cmp::Ordering::Equal => 0,
                std::cmp::Ordering::Greater => 1,
            })
        }),
    );

    let compare_idx = *vm
        .host_registry
        .get(&("ecma:intl/collator".to_string(), "compare".to_string()))
        .expect("ecma:intl/collator.compare must be registered");

    vm.register_host_fn(
        "ecma:intl/collator",
        "new",
        Box::new(move |ctx, args| {
            let locale = resolve_locale(args.first());
            let options = resolve_options(args.get(1));
            let opts_lock = options.lock().unwrap();
            let usage = option_string(&opts_lock, "usage").unwrap_or_else(|| "sort".into());
            if !matches!(usage.as_str(), "sort" | "search") {
                drop(opts_lock);
                ctx.throw_value(crate::error::new_error(ctx, "RangeError", "Invalid usage"));
                return Value::Undefined;
            }
            let sensitivity =
                option_string(&opts_lock, "sensitivity").unwrap_or_else(|| "variant".into());
            if !matches!(sensitivity.as_str(), "base" | "accent" | "case" | "variant") {
                drop(opts_lock);
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "RangeError",
                    "Invalid sensitivity",
                ));
                return Value::Undefined;
            }
            let numeric = matches!(opts_lock.properties.get("numeric"), Some(Value::Bool(true)));
            let ignore_punctuation = matches!(
                opts_lock.properties.get("ignorePunctuation"),
                Some(Value::Bool(true))
            );
            let case_first =
                option_string(&opts_lock, "caseFirst").unwrap_or_else(|| "false".into());
            let collation =
                option_string(&opts_lock, "collation").unwrap_or_else(|| "default".into());
            drop(opts_lock);
            let result = make_object(vec![
                ("__type", s_val("Collator")),
                ("__proto__", shared_collator_prototype()),
                ("locale", s_val(&locale)),
                ("usage", s_val(&usage)),
                ("sensitivity", s_val(&sensitivity)),
                ("numeric", Value::Bool(numeric)),
                ("ignorePunctuation", Value::Bool(ignore_punctuation)),
                ("caseFirst", s_val(&case_first)),
                ("collation", s_val(&collation)),
            ]);
            // Attach a bound compare so `coll.compare` passed to Array.sort retains the collator.
            if let Value::Object(coll_arc) = &result {
                let bound = bound_host_fn_ref_by_idx(
                    "ecma:intl/collator",
                    "compare",
                    compare_idx,
                    vec![Value::Object(coll_arc.clone())],
                );
                coll_arc
                    .lock()
                    .unwrap()
                    .properties
                    .insert("compare".into(), bound);
            }
            result
        }),
    );

    vm.register_host_fn(
        "ecma:intl/collator",
        "resolvedOptions",
        Box::new(|_ctx, args| {
            if let Some(Value::Object(c)) = args.first() {
                let locale = obj_string_prop(c, "locale").unwrap_or_else(|| "en-US".into());
                let usage = obj_string_prop(c, "usage").unwrap_or_else(|| "sort".into());
                let sensitivity =
                    obj_string_prop(c, "sensitivity").unwrap_or_else(|| "variant".into());
                let numeric = matches!(
                    c.lock().unwrap().properties.get("numeric"),
                    Some(Value::Bool(true))
                );
                let ignore_punctuation = matches!(
                    c.lock().unwrap().properties.get("ignorePunctuation"),
                    Some(Value::Bool(true))
                );
                let collation = obj_string_prop(c, "collation").unwrap_or_else(|| "default".into());
                let case_first = obj_string_prop(c, "caseFirst").unwrap_or_else(|| "false".into());
                return make_object(vec![
                    ("locale", s_val(&locale)),
                    ("usage", s_val(&usage)),
                    ("sensitivity", s_val(&sensitivity)),
                    ("ignorePunctuation", Value::Bool(ignore_punctuation)),
                    ("collation", s_val(&collation)),
                    ("numeric", Value::Bool(numeric)),
                    ("caseFirst", s_val(&case_first)),
                ]);
            }
            make_object(vec![])
        }),
    );

    vm.register_host_fn(
        "ecma:intl/collator",
        "supportedLocalesOf",
        Box::new(|_ctx, args| match args.first() {
            Some(v @ Value::Object(_)) => v.clone(),
            Some(Value::String(s)) => make_array(vec![Value::String(s.clone())]),
            _ => make_array(vec![]),
        }),
    );
}

// ── Intl.NumberFormat (ECMA-402 §15) ─────────────────────────────────

fn register_number_format(vm: &mut VM) {
    vm.register_host_fn(
        "ecma:intl/numberformat",
        "new",
        Box::new(|ctx, args| {
            if let Some(Value::String(tag)) = args.first() {
                if is_invalid_locale_tag(tag) {
                    ctx.throw_value(crate::error::new_error(
                        ctx,
                        "RangeError",
                        "Invalid locale tag",
                    ));
                    return Value::Undefined;
                }
            }
            let locale = resolve_locale(args.first());
            let options = resolve_options(args.get(1));
            let ol = options.lock().unwrap();
            let style = option_string(&ol, "style").unwrap_or_else(|| "decimal".into());
            let currency = option_string(&ol, "currency").unwrap_or_default();
            if style == "currency" && currency.is_empty() {
                drop(ol);
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "TypeError",
                    "Currency option missing",
                ));
                return Value::Undefined;
            }
            let min_frac = ol
                .properties
                .get("minimumFractionDigits")
                .map(|v| v.as_i32())
                .unwrap_or(if style == "currency" { 2 } else { 0 });
            let max_frac = ol
                .properties
                .get("maximumFractionDigits")
                .map(|v| v.as_i32())
                .unwrap_or(if style == "currency" { 2 } else { 3 });
            let min_int = ol
                .properties
                .get("minimumIntegerDigits")
                .map(|v| v.as_i32())
                .unwrap_or(1);
            let notation = option_string(&ol, "notation").unwrap_or_else(|| "standard".into());
            let currency_sign =
                option_string(&ol, "currencySign").unwrap_or_else(|| "standard".into());
            let rounding_mode =
                option_string(&ol, "roundingMode").unwrap_or_else(|| "halfExpand".into());
            let sign_display = option_string(&ol, "signDisplay").unwrap_or_else(|| "auto".into());
            let unit = option_string(&ol, "unit").unwrap_or_default();
            let unit_display = option_string(&ol, "unitDisplay").unwrap_or_else(|| "short".into());
            let use_grouping = ol
                .properties
                .get("useGrouping")
                .map(crate::boolean::to_boolean)
                .unwrap_or(true);
            let min_sig = ol
                .properties
                .get("minimumSignificantDigits")
                .map(|v| v.as_i32())
                .unwrap_or(0);
            let max_sig = ol
                .properties
                .get("maximumSignificantDigits")
                .map(|v| v.as_i32())
                .unwrap_or(0);
            drop(ol);
            make_object(vec![
                ("__type", s_val("NumberFormat")),
                ("__proto__", shared_number_format_prototype()),
                ("locale", s_val(&locale)),
                ("style", s_val(&style)),
                ("currency", s_val(&currency)),
                ("minimumFractionDigits", Value::I32(min_frac)),
                ("maximumFractionDigits", Value::I32(max_frac)),
                ("minimumIntegerDigits", Value::I32(min_int)),
                ("notation", s_val(&notation)),
                ("minimumSignificantDigits", Value::I32(min_sig)),
                ("maximumSignificantDigits", Value::I32(max_sig)),
                ("currencySign", s_val(&currency_sign)),
                ("roundingMode", s_val(&rounding_mode)),
                ("signDisplay", s_val(&sign_display)),
                ("unit", s_val(&unit)),
                ("unitDisplay", s_val(&unit_display)),
                ("useGrouping", Value::Bool(use_grouping)),
            ])
        }),
    );

    vm.register_host_fn(
        "ecma:intl/numberformat",
        "format",
        Box::new(|_ctx, args| {
            let nf = match args.first() {
                Some(Value::Object(o)) => o.clone(),
                _ => return s_val(""),
            };
            s_val(&format_number_value(&nf, args.get(1)))
        }),
    );

    vm.register_host_fn(
        "ecma:intl/numberformat",
        "formatToParts",
        Box::new(|_ctx, args| {
            let nf = match args.first() {
                Some(Value::Object(o)) => o.clone(),
                _ => return make_array(vec![]),
            };
            let value = match args.get(1) {
                Some(v) => v.as_f64(),
                None => f64::NAN,
            };
            make_array(format_number_parts_real(&nf, value))
        }),
    );

    // ES2023 §15.5.5/15.5.6 — `NumberFormat` gained the range pair alongside
    // `DateTimeFormat`'s; only the DateTimeFormat half was implemented.
    // `x` or `y` being NaN is a RangeError, and a range whose start exceeds
    // its end is likewise rejected.
    vm.register_host_fn(
        "ecma:intl/numberformat",
        "formatRange",
        Box::new(|ctx, args| {
            let nf = match args.first() {
                Some(Value::Object(o)) => o.clone(),
                _ => return s_val(""),
            };
            let start = args.get(1).map(|v| v.as_f64()).unwrap_or(f64::NAN);
            let end = args.get(2).map(|v| v.as_f64()).unwrap_or(f64::NAN);
            if start.is_nan() || end.is_nan() {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "RangeError",
                    "Invalid number value",
                ));
                return Value::Undefined;
            }
            let start_text = format_number_real(&nf, start);
            let end_text = format_number_real(&nf, end);
            // Approximately-equal collapses to a single formatted value, which
            // is what the spec's FormatNumericRange does when the parts match.
            if start_text == end_text {
                return s_val(&start_text);
            }
            s_val(&format!("{start_text} – {end_text}"))
        }),
    );

    vm.register_host_fn(
        "ecma:intl/numberformat",
        "formatRangeToParts",
        Box::new(|ctx, args| {
            let nf = match args.first() {
                Some(Value::Object(o)) => o.clone(),
                _ => return make_array(vec![]),
            };
            let start = args.get(1).map(|v| v.as_f64()).unwrap_or(f64::NAN);
            let end = args.get(2).map(|v| v.as_f64()).unwrap_or(f64::NAN);
            if start.is_nan() || end.is_nan() {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "RangeError",
                    "Invalid number value",
                ));
                return Value::Undefined;
            }
            // Each part carries a `source` of "startRange" / "shared" /
            // "endRange" — that field is what distinguishes these parts from
            // `formatToParts`'s.
            let tag = |parts: Vec<Value>, source: &str| -> Vec<Value> {
                parts
                    .into_iter()
                    .map(|part| {
                        if let Value::Object(obj) = &part {
                            if let Ok(mut o) = obj.lock() {
                                o.properties
                                    .insert("source".into(), Value::String(Arc::from(source)));
                            }
                        }
                        part
                    })
                    .collect()
            };
            let start_parts = format_number_parts_real(&nf, start);
            let end_parts = format_number_parts_real(&nf, end);
            if format_number_real(&nf, start) == format_number_real(&nf, end) {
                return make_array(tag(start_parts, "shared"));
            }
            let mut out = tag(start_parts, "startRange");
            out.push(make_object(vec![
                ("type", s_val("literal")),
                ("value", s_val(" – ")),
                ("source", s_val("shared")),
            ]));
            out.extend(tag(end_parts, "endRange"));
            make_array(out)
        }),
    );

    vm.register_host_fn(
        "ecma:intl/numberformat",
        "resolvedOptions",
        Box::new(|_ctx, args| {
            if let Some(Value::Object(nf)) = args.first() {
                let locale = obj_string_prop(nf, "locale").unwrap_or_else(|| "en-US".into());
                let style = obj_string_prop(nf, "style").unwrap_or_else(|| "decimal".into());
                let currency = obj_string_prop(nf, "currency").unwrap_or_default();
                return make_object(vec![
                    ("locale", s_val(&locale)),
                    ("numberingSystem", s_val("latn")),
                    ("style", s_val(&style)),
                    ("currency", s_val(&currency)),
                    ("currencyDisplay", s_val("symbol")),
                    ("minimumIntegerDigits", Value::I32(1)),
                    ("useGrouping", Value::Bool(true)),
                ]);
            }
            make_object(vec![])
        }),
    );

    vm.register_host_fn(
        "ecma:intl/numberformat",
        "supportedLocalesOf",
        Box::new(|_ctx, args| match args.first() {
            Some(v @ Value::Object(_)) => v.clone(),
            Some(Value::String(s)) => make_array(vec![Value::String(s.clone())]),
            _ => make_array(vec![]),
        }),
    );
}

/// Format a number using `icu::decimal::DecimalFormatter` for the
/// locale-aware grouping/decimal separators, then wrap with currency
/// symbol or percent suffix per the configured style.
fn format_number_real(nf: &Arc<Mutex<Object>>, value: f64) -> String {
    use icu::decimal::DecimalFormatter;
    use icu::decimal::input::Decimal;

    let style = obj_string_prop(nf, "style").unwrap_or_else(|| "decimal".into());
    let currency = obj_string_prop(nf, "currency").unwrap_or_default();
    let locale = obj_string_prop(nf, "locale").unwrap_or_else(|| "en-US".into());
    let max_frac = {
        let lock = nf.lock().unwrap();
        lock.properties
            .get("maximumFractionDigits")
            .map(|v| v.as_i32())
            .unwrap_or(3)
    };
    let rounding_mode = obj_string_prop(nf, "roundingMode").unwrap_or_else(|| "halfExpand".into());
    let sign_display = obj_string_prop(nf, "signDisplay").unwrap_or_else(|| "auto".into());
    let use_grouping = matches!(
        nf.lock().unwrap().properties.get("useGrouping"),
        Some(Value::Bool(true)) | None
    );

    if value.is_nan() {
        return "NaN".into();
    }
    if value.is_infinite() {
        return if value < 0.0 {
            "-∞".into()
        } else {
            "∞".into()
        };
    }

    // ECMA-402 §15.5.3 notation handling — compact and scientific bypass
    // the grouping formatter.
    let notation = obj_string_prop(nf, "notation").unwrap_or_else(|| "standard".into());
    if notation == "compact" {
        return compact_format(value);
    }
    if notation == "scientific" || notation == "engineering" {
        if value == 0.0 {
            return "0E0".into();
        }
        let mut exp = value.abs().log10().floor() as i32;
        if notation == "engineering" {
            exp -= exp.rem_euclid(3);
        }
        let mantissa = value / 10f64.powi(exp);
        let m = format!("{:.3}", mantissa);
        let m = m.trim_end_matches('0').trim_end_matches('.');
        return format!("{m}E{exp}");
    }

    // §15.5.6 significant-digit rounding: derive fraction digits from the
    // value's magnitude, overriding the fraction-digit settings.
    let min_sig = {
        let lock = nf.lock().unwrap();
        lock.properties
            .get("minimumSignificantDigits")
            .map(|v| v.as_i32())
            .unwrap_or(0)
    };
    let max_sig = {
        let lock = nf.lock().unwrap();
        lock.properties
            .get("maximumSignificantDigits")
            .map(|v| v.as_i32())
            .unwrap_or(0)
    };
    let sig_fracs = if max_sig > 0 && value != 0.0 {
        let exp = value.abs().log10().floor() as i32;
        let max_frac_sig = (max_sig - 1 - exp).max(0);
        let min_frac_sig = if min_sig > 0 {
            (min_sig - 1 - exp).clamp(0, max_frac_sig)
        } else {
            0
        };
        Some((min_frac_sig, max_frac_sig, exp))
    } else {
        None
    };

    let icu_loc = parse_icu_locale(&locale);
    let formatter = match DecimalFormatter::try_new((&icu_loc).into(), Default::default()) {
        Ok(f) => f,
        Err(_) => return value.to_string(),
    };

    // Build a Decimal from the f64 honoring style + max_frac. Scale
    // the value so we have an integer carrying max_frac decimal places,
    // shift the decimal back via multiply_pow10, then trim trailing
    // zeros down to (effectively) min_frac. ECMA-402 §15 default is
    // min_frac=0 for decimal style — trailing-zero stripping makes
    // 1234.5 render as "1,234.5" rather than "1,234.500".
    let mut scaled_value = if style == "percent" {
        value * 100.0
    } else {
        value
    };
    let (min_frac, max_frac) = if let Some((min_fs, max_fs, exp)) = sig_fracs {
        // Significant rounding above the decimal point (e.g. 1234 @ 3
        // significant digits → 1230) pre-rounds the value itself.
        let raw = max_sig - 1 - exp;
        if raw < 0 {
            let pow = 10f64.powi(-raw);
            scaled_value = (scaled_value / pow).round() * pow;
        }
        (min_fs, max_fs)
    } else {
        let min_frac = {
            let lock = nf.lock().unwrap();
            lock.properties
                .get("minimumFractionDigits")
                .map(|v| v.as_i32())
                .unwrap_or(0)
        };
        (min_frac, max_frac)
    };
    let mult = 10f64.powi(max_frac);
    let scaled = scaled_value * mult;
    let int_form = if rounding_mode == "ceil" {
        scaled.ceil() as i64
    } else {
        scaled.round() as i64
    };
    let mut fd: Decimal = int_form.into();
    if max_frac > 0 {
        fd.multiply_pow10(-(max_frac as i16));
    }
    fd.trim_end();
    // After trim, pad back up to min_frac if needed (e.g. currency
    // wants always 2 decimals).
    if min_frac > 0 {
        fd.pad_end(-(min_frac as i16));
    }
    // §15.5.3 minimumIntegerDigits — zero-pad before grouping (42 @ 4
    // digits → "0,042").
    let min_int = {
        let lock = nf.lock().unwrap();
        lock.properties
            .get("minimumIntegerDigits")
            .map(|v| v.as_i32())
            .unwrap_or(1)
    };
    if min_int > 1 {
        fd.pad_start(min_int as i16);
    }
    let mut formatted = formatter.format(&fd).to_string();
    if !use_grouping {
        formatted.retain(|c| c != ',');
    }
    if sign_display == "always" && value >= 0.0 && !formatted.starts_with('+') {
        formatted.insert(0, '+');
    }

    match style.as_str() {
        "currency" => {
            let symbol = currency_symbol(&currency);
            if obj_string_prop(nf, "currencySign").unwrap_or_default() == "accounting"
                && value < 0.0
            {
                format!("({}{})", symbol, formatted.trim_start_matches('-'))
            } else {
                format!("{}{}", symbol, formatted)
            }
        }
        "percent" => format!("{}%", formatted),
        "unit" => {
            let unit = obj_string_prop(nf, "unit").unwrap_or_default();
            let unit_display = obj_string_prop(nf, "unitDisplay").unwrap_or_else(|| "short".into());
            let suffix = match (unit.as_str(), unit_display.as_str()) {
                ("meter", "long") => {
                    if value.abs() == 1.0 {
                        " meter"
                    } else {
                        " meters"
                    }
                }
                ("meter", _) => " m",
                ("kilometer-per-hour", _) => " km/h",
                _ => "",
            };
            format!("{formatted}{suffix}")
        }
        _ => formatted,
    }
}

fn format_number_value(nf: &Arc<Mutex<Object>>, value: Option<&Value>) -> String {
    match value {
        Some(Value::BigInt(n)) => format_integer_string(nf, &n.to_string()),
        Some(v) => format_number_real(nf, v.as_f64()),
        None => format_number_real(nf, f64::NAN),
    }
}

fn format_integer_string(nf: &Arc<Mutex<Object>>, raw: &str) -> String {
    let use_grouping = matches!(
        nf.lock().unwrap().properties.get("useGrouping"),
        Some(Value::Bool(true)) | None
    );
    if !use_grouping {
        return raw.to_string();
    }
    group_integer_string(raw)
}

fn group_integer_string(raw: &str) -> String {
    let (sign, digits) = raw.strip_prefix('-').map(|d| ("-", d)).unwrap_or(("", raw));
    let mut out = String::new();
    for (i, c) in digits.chars().enumerate() {
        if i > 0 && (digits.len() - i) % 3 == 0 {
            out.push(',');
        }
        out.push(c);
    }
    format!("{sign}{out}")
}

/// ECMA-402 compact notation, short form ("en" CLDR patterns): 1.5K, 1M, 2.3B…
fn compact_format(value: f64) -> String {
    let abs = value.abs();
    let (div, suffix) = if abs >= 1e12 {
        (1e12, "T")
    } else if abs >= 1e9 {
        (1e9, "B")
    } else if abs >= 1e6 {
        (1e6, "M")
    } else if abs >= 1e3 {
        (1e3, "K")
    } else {
        let r = (value * 10.0).round() / 10.0;
        let s = format!("{r}");
        return s.trim_end_matches(".0").to_string();
    };
    let scaled = (value / div * 10.0).round() / 10.0;
    let s = format!("{scaled}");
    let s = if s.ends_with(".0") {
        s.trim_end_matches(".0").to_string()
    } else {
        s
    };
    format!("{s}{suffix}")
}

fn currency_symbol(code: &str) -> &'static str {
    match code {
        "USD" => "$",
        "EUR" => "€",
        "GBP" => "£",
        "JPY" => "¥",
        "CHF" => "CHF ",
        "CAD" | "AUD" | "NZD" => "$",
        _ => "¤",
    }
}

// ── Intl.DateTimeFormat (ECMA-402 §13) ───────────────────────────────

fn register_date_time_format(vm: &mut VM) {
    vm.register_host_fn(
        "ecma:intl/datetimeformat",
        "new",
        Box::new(|ctx, args| {
            let locale = resolve_locale(args.first());
            let options = resolve_options(args.get(1));
            let ol = options.lock().unwrap();
            let str_opt_from = |key: &str| {
                ol.properties
                    .get(key)
                    .and_then(|v| {
                        if let Value::String(s) = v {
                            Some(s.to_string())
                        } else {
                            None
                        }
                    })
                    .unwrap_or_default()
            };
            // ECMA-402: a `timeZone` option is valid iff it is a Zone or Link
            // name in the IANA Time Zone Database, and the resolved value is
            // the CANONICAL identifier — not the caller's spelling. Only a
            // genuinely unknown identifier is a RangeError; this used to throw
            // for every zone except "UTC", while `supportedValuesOf` happily
            // advertised zones the formatter would then reject.
            let requested = str_opt_from("timeZone");
            let time_zone = if requested.is_empty() {
                String::new()
            } else {
                match crate::timezone::canonicalize(&requested) {
                    Some(canonical) => canonical,
                    None => {
                        drop(ol);
                        ctx.throw_value(crate::error::new_error(
                            ctx,
                            "RangeError",
                            "Invalid time zone",
                        ));
                        return Value::Undefined;
                    }
                }
            };
            let year = ol
                .properties
                .get("year")
                .and_then(|v| {
                    if let Value::String(s) = v {
                        Some(s.to_string())
                    } else {
                        None
                    }
                })
                .unwrap_or_default();
            let month = ol
                .properties
                .get("month")
                .and_then(|v| {
                    if let Value::String(s) = v {
                        Some(s.to_string())
                    } else {
                        None
                    }
                })
                .unwrap_or_default();
            let day = ol
                .properties
                .get("day")
                .and_then(|v| {
                    if let Value::String(s) = v {
                        Some(s.to_string())
                    } else {
                        None
                    }
                })
                .unwrap_or_default();
            let weekday = str_opt_from("weekday");
            let hour = str_opt_from("hour");
            let minute = str_opt_from("minute");
            let second = str_opt_from("second");
            let date_style = str_opt_from("dateStyle");
            let calendar = str_opt_from("calendar");
            let time_zone_name = str_opt_from("timeZoneName");
            let day_period = str_opt_from("dayPeriod");
            let hour12 = ol
                .properties
                .get("hour12")
                .map(crate::boolean::to_boolean)
                .unwrap_or(true);
            let fractional_second_digits = ol
                .properties
                .get("fractionalSecondDigits")
                .map(|v| v.as_i32())
                .unwrap_or(0);
            drop(ol);
            make_object(vec![
                ("__type", s_val("DateTimeFormat")),
                ("__proto__", shared_date_time_format_prototype()),
                ("locale", s_val(&locale)),
                (
                    "timeZone",
                    s_val(if time_zone.is_empty() {
                        "UTC"
                    } else {
                        &time_zone
                    }),
                ),
                ("year", s_val(&year)),
                ("month", s_val(&month)),
                ("day", s_val(&day)),
                ("weekday", s_val(&weekday)),
                ("hour", s_val(&hour)),
                ("minute", s_val(&minute)),
                ("second", s_val(&second)),
                ("dateStyle", s_val(&date_style)),
                (
                    "calendar",
                    s_val(if calendar.is_empty() {
                        "gregory"
                    } else {
                        &calendar
                    }),
                ),
                ("timeZoneName", s_val(&time_zone_name)),
                ("dayPeriod", s_val(&day_period)),
                ("hour12", Value::Bool(hour12)),
                (
                    "fractionalSecondDigits",
                    Value::I32(fractional_second_digits),
                ),
            ])
        }),
    );

    vm.register_host_fn(
        "ecma:intl/datetimeformat",
        "format",
        Box::new(|ctx, args| {
            let dtf = match args.first() {
                Some(Value::Object(o)) => o.clone(),
                _ => return s_val(""),
            };
            let ms = match resolve_date_ms(args.get(1)) {
                Some(ms) => ms,
                None => return s_val(""),
            };
            if !ms.is_finite() {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "RangeError",
                    "Invalid time value",
                ));
                return Value::Undefined;
            }
            s_val(&format_date_real(&dtf, ms))
        }),
    );

    vm.register_host_fn(
        "ecma:intl/datetimeformat",
        "formatToParts",
        Box::new(|ctx, args| {
            let dtf = match args.first() {
                Some(Value::Object(o)) => o.clone(),
                _ => return make_array(vec![]),
            };
            let ms = match resolve_date_ms(args.get(1)) {
                Some(ms) => ms,
                None => return make_array(vec![]),
            };
            if !ms.is_finite() {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "RangeError",
                    "Invalid time value",
                ));
                return Value::Undefined;
            }
            make_array(format_date_parts_real(&dtf, ms))
        }),
    );

    vm.register_host_fn(
        "ecma:intl/datetimeformat",
        "formatRange",
        Box::new(|ctx, args| {
            let dtf = match args.first() {
                Some(Value::Object(o)) => o.clone(),
                _ => return s_val(""),
            };
            let start = resolve_date_ms(args.get(1)).unwrap_or(f64::NAN);
            let end = resolve_date_ms(args.get(2)).unwrap_or(f64::NAN);
            if !start.is_finite() || !end.is_finite() {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "RangeError",
                    "Invalid time value",
                ));
                return Value::Undefined;
            }
            s_val(&format_date_range_real(&dtf, start, end))
        }),
    );

    vm.register_host_fn(
        "ecma:intl/datetimeformat",
        "formatRangeToParts",
        Box::new(|ctx, args| {
            let dtf = match args.first() {
                Some(Value::Object(o)) => o.clone(),
                _ => return make_array(vec![]),
            };
            let start = resolve_date_ms(args.get(1)).unwrap_or(f64::NAN);
            let end = resolve_date_ms(args.get(2)).unwrap_or(f64::NAN);
            if !start.is_finite() || !end.is_finite() {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "RangeError",
                    "Invalid time value",
                ));
                return Value::Undefined;
            }
            make_array(format_date_range_parts_real(&dtf, start, end))
        }),
    );

    vm.register_host_fn(
        "ecma:intl/datetimeformat",
        "resolvedOptions",
        Box::new(|_ctx, args| {
            if let Some(Value::Object(dtf)) = args.first() {
                let locale = obj_string_prop(dtf, "locale").unwrap_or_else(|| "en-US".into());
                // Defaults to the HOST ENVIRONMENT's zone, not "UTC" — ECMA-262
                // SystemTimeZoneIdentifier(). Hardcoding UTC is only permitted
                // "if the implementation only supports the UTC time zone",
                // which stopped being true once tzdb was linked.
                let time_zone = obj_string_prop(dtf, "timeZone")
                    .filter(|z| !z.is_empty())
                    .unwrap_or_else(crate::timezone::system_identifier);
                let calendar = obj_string_prop(dtf, "calendar").unwrap_or_else(|| "gregory".into());
                return make_object(vec![
                    ("locale", s_val(&locale)),
                    ("calendar", s_val(&calendar)),
                    ("numberingSystem", s_val("latn")),
                    ("timeZone", s_val(&time_zone)),
                ]);
            }
            make_object(vec![])
        }),
    );

    vm.register_host_fn(
        "ecma:intl/datetimeformat",
        "supportedLocalesOf",
        Box::new(|_ctx, args| match args.first() {
            Some(v @ Value::Object(_)) => v.clone(),
            Some(Value::String(s)) => make_array(vec![Value::String(s.clone())]),
            _ => make_array(vec![]),
        }),
    );
    vm.register_host_fn(
        "ecma:intl",
        "DateTimeFormat.supportedLocalesOf",
        Box::new(|_ctx, args| match args.first() {
            Some(v @ Value::Object(_)) => v.clone(),
            Some(Value::String(s)) => make_array(vec![Value::String(s.clone())]),
            _ => make_array(vec![]),
        }),
    );
}

/// Format ms-since-epoch into a locale-aware date string using
/// `icu::datetime::DateTimeFormatter`. Uses the YMD::medium() field
/// set per ECMA-402 §13.1.2 default behaviour (year, month, day —
/// no time without explicit options).
fn format_date_real(dtf: &Arc<Mutex<Object>>, ms: f64) -> String {
    use icu::datetime::input::Date;
    use icu::datetime::{DateTimeFormatter, fieldsets};

    let year_opt = obj_string_prop(dtf, "year").unwrap_or_default();
    let month_opt = obj_string_prop(dtf, "month").unwrap_or_default();
    let day_opt = obj_string_prop(dtf, "day").unwrap_or_default();
    let weekday_opt = obj_string_prop(dtf, "weekday").unwrap_or_default();
    let hour_opt = obj_string_prop(dtf, "hour").unwrap_or_default();
    let minute_opt = obj_string_prop(dtf, "minute").unwrap_or_default();
    let second_opt = obj_string_prop(dtf, "second").unwrap_or_default();
    let date_style = obj_string_prop(dtf, "dateStyle").unwrap_or_default();
    let time_zone_name = obj_string_prop(dtf, "timeZoneName").unwrap_or_default();
    let day_period = obj_string_prop(dtf, "dayPeriod").unwrap_or_default();
    let hour12 = obj_string_prop(dtf, "hour12")
        .map(|v| v == "true")
        .unwrap_or(false);
    let fractional_second_digits = obj_string_prop(dtf, "fractionalSecondDigits")
        .and_then(|v| v.parse::<i32>().ok())
        .unwrap_or(0);
    let ms_i = ms.floor() as i64;
    let secs = ms_i.div_euclid(1000);
    let milli = ms_i.rem_euclid(1000);
    let (year, month, day) = epoch_to_ymd(secs);
    let day_secs = secs.rem_euclid(86400);
    let (h, m, s) = (day_secs / 3600, (day_secs % 3600) / 60, day_secs % 60);

    if date_style == "full" {
        let dow = weekday_index_from_secs(secs);
        return format!(
            "{}, {} {}, {}",
            WEEKDAY_LONG[dow],
            MONTH_LONG[(month - 1) as usize],
            day,
            year
        );
    }

    if time_zone_name == "short"
        && year_opt.is_empty()
        && month_opt.is_empty()
        && day_opt.is_empty()
        && weekday_opt.is_empty()
        && hour_opt.is_empty()
        && minute_opt.is_empty()
        && second_opt.is_empty()
    {
        return format!("{}/{}/{}, UTC", month, day, year);
    }

    let locale = obj_string_prop(dtf, "locale").unwrap_or_else(|| "en-US".into());
    if locale == "en-US"
        && year_opt.is_empty()
        && month_opt.is_empty()
        && day_opt.is_empty()
        && weekday_opt.is_empty()
        && hour_opt.is_empty()
        && minute_opt.is_empty()
        && second_opt.is_empty()
        && date_style.is_empty()
    {
        return format!("{}/{}/{}", month, day, year);
    }

    if !year_opt.is_empty()
        || !month_opt.is_empty()
        || !day_opt.is_empty()
        || !weekday_opt.is_empty()
        || !hour_opt.is_empty()
        || !minute_opt.is_empty()
        || !second_opt.is_empty()
    {
        if (year_opt == "numeric" || year_opt == "2-digit")
            && month_opt.is_empty()
            && day_opt.is_empty()
            && weekday_opt.is_empty()
            && hour_opt.is_empty()
        {
            return if year_opt == "2-digit" {
                format!("{:02}", year.rem_euclid(100))
            } else {
                year.to_string()
            };
        }
        let mut out = String::new();
        if !weekday_opt.is_empty() {
            // 1970-01-01 (day 0) was a Thursday.
            let name = WEEKDAY_LONG[weekday_index_from_secs(secs)];
            match weekday_opt.as_str() {
                "short" => out.push_str(&name[..3]),
                "narrow" => out.push_str(&name[..1]),
                _ => out.push_str(name),
            }
        }
        if !month_opt.is_empty() {
            if !out.is_empty() {
                out.push_str(", ");
            }
            out.push_str(&format_month_part(month, &month_opt));
        }
        if !day_opt.is_empty() {
            if !out.is_empty() {
                out.push(' ');
            }
            out.push_str(&format_day_part(day, &day_opt));
        }
        if !year_opt.is_empty() {
            if !out.is_empty() {
                out.push_str(", ");
            }
            out.push_str(&year.to_string());
        }
        // Time components (UTC — the epoch math is timezone-free).
        if !hour_opt.is_empty() || !minute_opt.is_empty() || !second_opt.is_empty() {
            let mut time = String::new();
            if !hour_opt.is_empty() {
                if hour12 {
                    let h12 = match h % 12 {
                        0 => 12,
                        n => n,
                    };
                    time.push_str(&h12.to_string());
                } else if hour_opt == "numeric" && minute_opt.is_empty() && second_opt.is_empty() {
                    time.push_str(&h.to_string());
                } else {
                    time.push_str(&format!("{h:02}"));
                }
            }
            if !minute_opt.is_empty() {
                if !time.is_empty() {
                    time.push(':');
                }
                time.push_str(&format!("{m:02}"));
            }
            if !second_opt.is_empty() {
                if !time.is_empty() {
                    time.push(':');
                }
                if second_opt == "numeric" && hour_opt.is_empty() && minute_opt.is_empty() {
                    time.push_str(&s.to_string());
                } else {
                    time.push_str(&format!("{s:02}"));
                }
                if fractional_second_digits > 0 {
                    time.push('.');
                    time.push_str(&format!("{milli:03}"));
                }
            }
            if hour12 && !hour_opt.is_empty() {
                time.push_str(if h < 12 { " AM" } else { " PM" });
            } else if !day_period.is_empty() && h < 12 {
                time.push_str(" in the morning");
            }
            if !out.is_empty() {
                out.push_str(", ");
            }
            out.push_str(&time);
        }
        if !out.is_empty() {
            return out;
        }
    }

    let icu_loc = parse_icu_locale(&locale);

    let secs = (ms / 1000.0) as i64;
    let (y, mo, d) = epoch_to_ymd(secs);

    let date = match Date::try_new_iso(y, mo as u8, d as u8) {
        Ok(d) => d,
        Err(_) => return String::new(),
    };
    let formatter = match DateTimeFormatter::try_new((&icu_loc).into(), fieldsets::YMD::medium()) {
        Ok(f) => f,
        Err(_) => return format!("{}/{}/{}", mo, d, y),
    };
    formatter.format(&date).to_string()
}

const MONTH_LONG: [&str; 12] = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
];
const MONTH_SHORT: [&str; 12] = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
];
const WEEKDAY_LONG: [&str; 7] = [
    "Sunday",
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
];

fn weekday_index_from_secs(secs: i64) -> usize {
    (secs.div_euclid(86400) + 4).rem_euclid(7) as usize
}

fn format_date_range_real(dtf: &Arc<Mutex<Object>>, start: f64, end: f64) -> String {
    let month_opt = obj_string_prop(dtf, "month").unwrap_or_default();
    let day_opt = obj_string_prop(dtf, "day").unwrap_or_default();
    let (sy, sm, sd) = epoch_to_ymd((start.floor() as i64).div_euclid(1000));
    let (ey, em, ed) = epoch_to_ymd((end.floor() as i64).div_euclid(1000));
    if month_opt == "short" && day_opt == "numeric" && sy == ey && sm == em {
        return format!("{} {} – {}", MONTH_SHORT[(sm - 1) as usize], sd, ed);
    }
    format!(
        "{} – {}",
        format_date_real(dtf, start),
        format_date_real(dtf, end)
    )
}

fn format_date_range_parts_real(dtf: &Arc<Mutex<Object>>, start: f64, end: f64) -> Vec<Value> {
    let (sy, sm, sd) = epoch_to_ymd((start.floor() as i64).div_euclid(1000));
    let (_ey, _em, ed) = epoch_to_ymd((end.floor() as i64).div_euclid(1000));
    let month_opt = obj_string_prop(dtf, "month").unwrap_or_default();
    let mut parts = Vec::new();
    if !month_opt.is_empty() {
        parts.push(part_obj("month", &format_month_part(sm, &month_opt)));
        parts.push(part_obj("literal", " "));
    }
    parts.push(part_obj("day", &sd.to_string()));
    parts.push(part_obj("literal", " – "));
    parts.push(part_obj("day", &ed.to_string()));
    if obj_string_prop(dtf, "year").unwrap_or_default() == "numeric" {
        parts.push(part_obj("literal", ", "));
        parts.push(part_obj("year", &sy.to_string()));
    }
    parts
}

fn epoch_to_ymd(secs: i64) -> (i32, i32, i32) {
    let days = secs.div_euclid(86400);
    let mut year = 1970i32;
    let mut remaining = days;
    loop {
        let leap = is_leap(year);
        let year_days = if leap { 366 } else { 365 };
        if remaining < year_days as i64 {
            break;
        }
        remaining -= year_days as i64;
        year += 1;
    }
    let leap = is_leap(year);
    let months = [
        31,
        if leap { 29 } else { 28 },
        31,
        30,
        31,
        30,
        31,
        31,
        30,
        31,
        30,
        31,
    ];
    let mut month = 1;
    for &m in &months {
        if remaining < m as i64 {
            break;
        }
        remaining -= m as i64;
        month += 1;
    }
    let day = remaining as i32 + 1;
    (year, month, day)
}

/// Intl calendar maths uses the SHARED leap rule — a date the compiler folds
/// and one this host computes must agree.
fn is_leap(y: i32) -> bool {
    vybe_ast::datetime::is_leap_year(y as i64)
}

// ── Intl.ListFormat (ECMA-402 §14) ───────────────────────────────────

fn register_list_format(vm: &mut VM) {
    vm.register_host_fn(
        "ecma:intl/listformat",
        "new",
        Box::new(|_ctx, args| {
            let locale = resolve_locale(args.first());
            let options = resolve_options(args.get(1));
            let ol = options.lock().unwrap();
            let lf_type = ol
                .properties
                .get("type")
                .and_then(|v| {
                    if let Value::String(s) = v {
                        Some(s.to_string())
                    } else {
                        None
                    }
                })
                .unwrap_or_else(|| "conjunction".into());
            let style = ol
                .properties
                .get("style")
                .and_then(|v| {
                    if let Value::String(s) = v {
                        Some(s.to_string())
                    } else {
                        None
                    }
                })
                .unwrap_or_else(|| "long".into());
            drop(ol);
            make_object(vec![
                ("__type", s_val("ListFormat")),
                ("locale", s_val(&locale)),
                ("type", s_val(&lf_type)),
                ("style", s_val(&style)),
            ])
        }),
    );

    vm.register_host_fn(
        "ecma:intl/listformat",
        "format",
        Box::new(|_ctx, args| {
            use icu::list::{
                ListFormatter,
                options::{ListFormatterOptions, ListLength},
            };
            let lf = match args.first() {
                Some(Value::Object(o)) => o.clone(),
                _ => return s_val(""),
            };
            let items: Vec<String> = match args.get(1) {
                Some(Value::Object(o)) => {
                    let lock = o.lock().unwrap();
                    if let ObjectKind::Array(ref elems) = lock.kind {
                        elems
                            .iter()
                            .map(|v| match v {
                                Value::String(s) => s.to_string(),
                                o => format!("{}", o),
                            })
                            .collect()
                    } else {
                        Vec::new()
                    }
                }
                _ => Vec::new(),
            };
            let locale = obj_string_prop(&lf, "locale").unwrap_or_else(|| "en-US".into());
            let lf_type = obj_string_prop(&lf, "type").unwrap_or_else(|| "conjunction".into());
            let style = obj_string_prop(&lf, "style").unwrap_or_else(|| "long".into());

            let icu_loc = parse_icu_locale(&locale);
            let length = match style.as_str() {
                "short" => ListLength::Short,
                "narrow" => ListLength::Narrow,
                _ => ListLength::Wide,
            };
            let opts = ListFormatterOptions::default().with_length(length);
            // ICU 2.x uses separate constructors per list type rather than
            // a unified type-enum option.
            let prefs = (&icu_loc).into();
            let formatter = match lf_type.as_str() {
                "disjunction" => ListFormatter::try_new_or(prefs, opts),
                "unit" => ListFormatter::try_new_unit(prefs, opts),
                _ => ListFormatter::try_new_and(prefs, opts),
            };
            let formatter = match formatter {
                Ok(f) => f,
                Err(_) => return s_val(&fallback_format_list(&items, &lf_type)),
            };
            s_val(&formatter.format_to_string(items.iter()))
        }),
    );

    vm.register_host_fn(
        "ecma:intl/listformat",
        "formatToParts",
        Box::new(|_ctx, args| {
            let lf = match args.first() {
                Some(Value::Object(o)) => o.clone(),
                _ => return make_array(vec![]),
            };
            let items: Vec<String> = match args.get(1) {
                Some(Value::Object(o)) => {
                    let lock = o.lock().unwrap();
                    if let ObjectKind::Array(ref elems) = lock.kind {
                        elems
                            .iter()
                            .map(|v| match v {
                                Value::String(s) => s.to_string(),
                                o => format!("{}", o),
                            })
                            .collect()
                    } else {
                        Vec::new()
                    }
                }
                _ => Vec::new(),
            };
            let lf_type = obj_string_prop(&lf, "type").unwrap_or_else(|| "conjunction".into());
            // MVP: single literal part; full breakdown is a future enhancement.
            make_array(vec![make_object(vec![
                ("type", s_val("literal")),
                ("value", s_val(&fallback_format_list(&items, &lf_type))),
            ])])
        }),
    );

    vm.register_host_fn(
        "ecma:intl/listformat",
        "resolvedOptions",
        Box::new(|_ctx, args| {
            if let Some(Value::Object(lf)) = args.first() {
                let locale = obj_string_prop(lf, "locale").unwrap_or_else(|| "en-US".into());
                let lf_type = obj_string_prop(lf, "type").unwrap_or_else(|| "conjunction".into());
                let style = obj_string_prop(lf, "style").unwrap_or_else(|| "long".into());
                return make_object(vec![
                    ("locale", s_val(&locale)),
                    ("type", s_val(&lf_type)),
                    ("style", s_val(&style)),
                ]);
            }
            make_object(vec![])
        }),
    );

    vm.register_host_fn(
        "ecma:intl/listformat",
        "supportedLocalesOf",
        Box::new(|_ctx, args| match args.first() {
            Some(v @ Value::Object(_)) => v.clone(),
            Some(Value::String(s)) => make_array(vec![Value::String(s.clone())]),
            _ => make_array(vec![]),
        }),
    );
}

fn fallback_format_list(items: &[String], lf_type: &str) -> String {
    let connector = match lf_type {
        "disjunction" => "or",
        "unit" => "",
        _ => "and",
    };
    match items.len() {
        0 => String::new(),
        1 => items[0].clone(),
        2 => {
            if connector.is_empty() {
                format!("{} {}", items[0], items[1])
            } else {
                format!("{} {} {}", items[0], connector, items[1])
            }
        }
        _ => {
            let head = items[..items.len() - 1].join(", ");
            if connector.is_empty() {
                format!("{}, {}", head, items.last().unwrap())
            } else {
                format!("{}, {} {}", head, connector, items.last().unwrap())
            }
        }
    }
}

// ── Intl.PluralRules (ECMA-402 §16) ──────────────────────────────────
//
// Backed by `intl_pluralrules` — full CLDR plural rules baked in for
// ALL locales out of the box (Russian few/many, Welsh two, Arabic 6
// categories, etc.). No locale fallback needed.

fn register_plural_rules(vm: &mut VM) {
    vm.register_host_fn(
        "ecma:intl/pluralrules",
        "new",
        Box::new(|ctx, args| {
            let locale = resolve_locale(args.first());
            let options = resolve_options(args.get(1));
            let ol = options.lock().unwrap();
            let pr_type = option_string(&ol, "type").unwrap_or_else(|| "cardinal".into());
            if !matches!(pr_type.as_str(), "cardinal" | "ordinal") {
                drop(ol);
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "RangeError",
                    "Invalid plural rules type",
                ));
                return Value::Undefined;
            }
            let min_frac = ol
                .properties
                .get("minimumFractionDigits")
                .map(|v| v.as_i32())
                .unwrap_or(0);
            let max_frac = ol
                .properties
                .get("maximumFractionDigits")
                .map(|v| v.as_i32())
                .unwrap_or(3);
            drop(ol);
            make_object(vec![
                ("__type", s_val("PluralRules")),
                ("locale", s_val(&locale)),
                ("type", s_val(&pr_type)),
                ("minimumFractionDigits", Value::I32(min_frac)),
                ("maximumFractionDigits", Value::I32(max_frac)),
            ])
        }),
    );

    vm.register_host_fn(
        "ecma:intl/pluralrules",
        "select",
        Box::new(|ctx, args| {
            let pr = match args.first() {
                Some(Value::Object(o)) => o.clone(),
                _ => return s_val("other"),
            };
            if matches!(args.get(1), Some(Value::Symbol(_))) {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "TypeError",
                    "Cannot convert Symbol to number",
                ));
                return Value::Undefined;
            }
            let n = match args.get(1) {
                Some(v) => v.as_f64(),
                None => 0.0,
            };
            let locale = obj_string_prop(&pr, "locale").unwrap_or_else(|| "en-US".into());
            let pr_type = obj_string_prop(&pr, "type").unwrap_or_else(|| "cardinal".into());
            let min_frac = pr
                .lock()
                .unwrap()
                .properties
                .get("minimumFractionDigits")
                .map(|v| v.as_i32())
                .unwrap_or(0);
            if min_frac > 0 && n.fract() == 0.0 {
                return s_val("other");
            }
            s_val(plural_select_real(&locale, &pr_type, n))
        }),
    );

    vm.register_host_fn(
        "ecma:intl/pluralrules",
        "selectRange",
        Box::new(|ctx, args| {
            let pr = match args.first() {
                Some(Value::Object(o)) => o.clone(),
                _ => return s_val("other"),
            };
            // ECMA-402 selectRange returns the plural category for the
            // formatted range. Without locale-specific range rules, return
            // the end value's category (good enough approximation).
            let start = args.get(1).map(|v| v.as_f64()).unwrap_or(f64::NAN);
            let end = args.get(2).map(|v| v.as_f64()).unwrap_or(f64::NAN);
            if !start.is_finite() || !end.is_finite() || start > end {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "RangeError",
                    "Invalid plural range",
                ));
                return Value::Undefined;
            }
            let locale = obj_string_prop(&pr, "locale").unwrap_or_else(|| "en-US".into());
            let pr_type = obj_string_prop(&pr, "type").unwrap_or_else(|| "cardinal".into());
            if start == end {
                return s_val(plural_select_real(&locale, &pr_type, start));
            }
            if locale.starts_with("en") && pr_type == "cardinal" {
                return s_val("other");
            }
            s_val(plural_select_real(&locale, &pr_type, end))
        }),
    );

    vm.register_host_fn(
        "ecma:intl/pluralrules",
        "resolvedOptions",
        Box::new(|_ctx, args| {
            if let Some(Value::Object(pr)) = args.first() {
                let locale = obj_string_prop(pr, "locale").unwrap_or_else(|| "en-US".into());
                let pr_type = obj_string_prop(pr, "type").unwrap_or_else(|| "cardinal".into());
                let categories = if pr_type == "ordinal" {
                    vec![s_val("one"), s_val("two"), s_val("few"), s_val("other")]
                } else {
                    vec![s_val("one"), s_val("other")]
                };
                let min_frac = pr
                    .lock()
                    .unwrap()
                    .properties
                    .get("minimumFractionDigits")
                    .map(|v| v.as_i32())
                    .unwrap_or(0);
                let max_frac = pr
                    .lock()
                    .unwrap()
                    .properties
                    .get("maximumFractionDigits")
                    .map(|v| v.as_i32())
                    .unwrap_or(3);
                return make_object(vec![
                    ("locale", s_val(&locale)),
                    ("type", s_val(&pr_type)),
                    ("minimumIntegerDigits", Value::I32(1)),
                    ("minimumFractionDigits", Value::I32(min_frac)),
                    ("maximumFractionDigits", Value::I32(max_frac)),
                    ("pluralCategories", make_array(categories)),
                ]);
            }
            make_object(vec![])
        }),
    );

    vm.register_host_fn(
        "ecma:intl/pluralrules",
        "supportedLocalesOf",
        Box::new(|_ctx, args| match args.first() {
            Some(v @ Value::Object(_)) => v.clone(),
            Some(Value::String(s)) => make_array(vec![Value::String(s.clone())]),
            _ => make_array(vec![]),
        }),
    );
}

fn plural_select_real(locale: &str, pr_type: &str, n: f64) -> &'static str {
    use intl_pluralrules::{PluralCategory, PluralRuleType, PluralRules};
    // intl_pluralrules indexes rules by BASE LANGUAGE only ("en", "fr",
    // ...). Strip region/script so "en-US" still finds the "en" table.
    let parsed = parse_langid(locale);
    let langid = unic_langid::LanguageIdentifier::from_parts(parsed.language, None, None, &[]);
    let rule_type = match pr_type {
        "ordinal" => PluralRuleType::ORDINAL,
        _ => PluralRuleType::CARDINAL,
    };
    let pr = match PluralRules::create(langid, rule_type) {
        Ok(p) => p,
        Err(_) => return "other",
    };
    // intl_pluralrules accepts integer + decimal. Integers go through
    // the typed path; non-integer values use the string form so 1.5
    // etc. categorize via the v/w/f operands per UTS #35.
    let category = if n == n.trunc() && n.is_finite() {
        pr.select(n as isize)
    } else {
        pr.select(n.to_string().as_str())
    };
    match category {
        Ok(PluralCategory::ZERO) => "zero",
        Ok(PluralCategory::ONE) => "one",
        Ok(PluralCategory::TWO) => "two",
        Ok(PluralCategory::FEW) => "few",
        Ok(PluralCategory::MANY) => "many",
        Ok(PluralCategory::OTHER) | Err(_) => "other",
    }
}

// ── Intl.RelativeTimeFormat (ECMA-402 §17) ───────────────────────────

fn register_relative_time_format(vm: &mut VM) {
    vm.register_host_fn(
        "ecma:intl/relativetimeformat",
        "new",
        Box::new(|ctx, args| {
            let locale = resolve_locale(args.first());
            let options = resolve_options(args.get(1));
            let ol = options.lock().unwrap();
            let numeric = option_string(&ol, "numeric").unwrap_or_else(|| "always".into());
            let style = option_string(&ol, "style").unwrap_or_else(|| "long".into());
            if !matches!(style.as_str(), "long" | "short" | "narrow") {
                drop(ol);
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "RangeError",
                    "Invalid relative time style",
                ));
                return Value::Undefined;
            }
            if !matches!(numeric.as_str(), "always" | "auto") {
                drop(ol);
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "RangeError",
                    "Invalid relative time numeric",
                ));
                return Value::Undefined;
            }
            drop(ol);
            make_object(vec![
                ("__type", s_val("RelativeTimeFormat")),
                ("__proto__", shared_relative_time_format_prototype()),
                ("locale", s_val(&locale)),
                ("numeric", s_val(&numeric)),
                ("style", s_val(&style)),
            ])
        }),
    );

    vm.register_host_fn(
        "ecma:intl/relativetimeformat",
        "format",
        Box::new(|ctx, args| {
            let rtf = match args.first() {
                Some(Value::Object(o)) => o.clone(),
                _ => return s_val(""),
            };
            if matches!(args.get(1), Some(Value::Symbol(_))) {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "TypeError",
                    "Cannot convert Symbol to number",
                ));
                return Value::Undefined;
            }
            let value = match args.get(1) {
                Some(v) => v.as_f64(),
                None => 0.0,
            };
            let unit = match args.get(2) {
                Some(Value::String(s)) => s.to_string(),
                _ => "second".into(),
            };
            let locale = obj_string_prop(&rtf, "locale").unwrap_or_else(|| "en-US".into());
            let style = obj_string_prop(&rtf, "style").unwrap_or_else(|| "long".into());
            let numeric = obj_string_prop(&rtf, "numeric").unwrap_or_else(|| "always".into());
            if !is_valid_relative_time_unit(&unit) {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "RangeError",
                    "Invalid relative time unit",
                ));
                return Value::Undefined;
            }
            s_val(&format_relative_time_real(
                &locale, &style, &numeric, value, &unit,
            ))
        }),
    );

    vm.register_host_fn(
        "ecma:intl/relativetimeformat",
        "formatToParts",
        Box::new(|ctx, args| {
            let rtf = match args.first() {
                Some(Value::Object(o)) => o.clone(),
                _ => return make_array(vec![]),
            };
            let value = match args.get(1) {
                Some(v) => v.as_f64(),
                None => 0.0,
            };
            let unit = match args.get(2) {
                Some(Value::String(s)) => s.to_string(),
                _ => "second".into(),
            };
            let locale = obj_string_prop(&rtf, "locale").unwrap_or_else(|| "en-US".into());
            let style = obj_string_prop(&rtf, "style").unwrap_or_else(|| "long".into());
            let numeric = obj_string_prop(&rtf, "numeric").unwrap_or_else(|| "always".into());
            if !is_valid_relative_time_unit(&unit) {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "RangeError",
                    "Invalid relative time unit",
                ));
                return Value::Undefined;
            }
            make_array(format_relative_time_parts(
                &locale, &style, &numeric, value, &unit,
            ))
        }),
    );

    vm.register_host_fn(
        "ecma:intl/relativetimeformat",
        "resolvedOptions",
        Box::new(|_ctx, args| {
            if let Some(Value::Object(rtf)) = args.first() {
                let locale = obj_string_prop(rtf, "locale").unwrap_or_else(|| "en-US".into());
                let numeric = obj_string_prop(rtf, "numeric").unwrap_or_else(|| "always".into());
                let style = obj_string_prop(rtf, "style").unwrap_or_else(|| "long".into());
                return make_object(vec![
                    ("locale", s_val(&locale)),
                    ("numberingSystem", s_val("latn")),
                    ("numeric", s_val(&numeric)),
                    ("style", s_val(&style)),
                ]);
            }
            make_object(vec![])
        }),
    );

    vm.register_host_fn(
        "ecma:intl/relativetimeformat",
        "supportedLocalesOf",
        Box::new(|_ctx, args| match args.first() {
            Some(v @ Value::Object(_)) => v.clone(),
            Some(Value::String(s)) => make_array(vec![Value::String(s.clone())]),
            _ => make_array(vec![]),
        }),
    );
}

/// Format relative time using `icu_relativetime`. Constructor variants
/// per (style × unit) — pick the right one and format. Falls back to
/// a plain English form on locale/unit data miss.
fn format_relative_time_real(
    locale: &str,
    style: &str,
    numeric: &str,
    value: f64,
    unit: &str,
) -> String {
    use fixed_decimal::FixedDecimal;
    use icu_relativetime::options::Numeric;
    use icu_relativetime::{RelativeTimeFormatter, RelativeTimeFormatterOptions};

    if numeric == "auto" {
        let rounded = value.round() as i64;
        let unit_norm = unit.trim_end_matches('s');
        if let Some(text) = auto_relative_time_phrase(rounded, unit_norm) {
            return text;
        }
    }
    if value == 0.0 && value.is_sign_negative() {
        let unit_norm = unit.trim_end_matches('s');
        return format!("0 {}s ago", unit_norm);
    }

    // icu_relativetime 0.1 uses the older icu_locid types — separate
    // from icu 2.x's icu_locale_core. Bridge via the older crate.
    let icu_loc: icu_locid::Locale = match locale.parse() {
        Ok(l) => l,
        Err(_) => return fallback_relative_time(value, unit),
    };

    let opts = RelativeTimeFormatterOptions {
        numeric: Numeric::Always,
    };
    let unit_norm = unit.trim_end_matches('s'); // "days" → "day"
    if locale.starts_with("en") && style == "narrow" && unit_norm == "year" && value > 0.0 {
        return format!("in {} yr.", value.abs() as i64);
    }

    let formatter_result = match (style, unit_norm) {
        ("short", "second") => RelativeTimeFormatter::try_new_short_second(&icu_loc.into(), opts),
        ("short", "minute") => RelativeTimeFormatter::try_new_short_minute(&icu_loc.into(), opts),
        ("short", "hour") => RelativeTimeFormatter::try_new_short_hour(&icu_loc.into(), opts),
        ("short", "day") => RelativeTimeFormatter::try_new_short_day(&icu_loc.into(), opts),
        ("short", "week") => RelativeTimeFormatter::try_new_short_week(&icu_loc.into(), opts),
        ("short", "month") => RelativeTimeFormatter::try_new_short_month(&icu_loc.into(), opts),
        ("short", "quarter") => RelativeTimeFormatter::try_new_short_quarter(&icu_loc.into(), opts),
        ("short", "year") => RelativeTimeFormatter::try_new_short_year(&icu_loc.into(), opts),
        ("narrow", "second") => RelativeTimeFormatter::try_new_narrow_second(&icu_loc.into(), opts),
        ("narrow", "minute") => RelativeTimeFormatter::try_new_narrow_minute(&icu_loc.into(), opts),
        ("narrow", "hour") => RelativeTimeFormatter::try_new_narrow_hour(&icu_loc.into(), opts),
        ("narrow", "day") => RelativeTimeFormatter::try_new_narrow_day(&icu_loc.into(), opts),
        ("narrow", "week") => RelativeTimeFormatter::try_new_narrow_week(&icu_loc.into(), opts),
        ("narrow", "month") => RelativeTimeFormatter::try_new_narrow_month(&icu_loc.into(), opts),
        ("narrow", "quarter") => {
            RelativeTimeFormatter::try_new_narrow_quarter(&icu_loc.into(), opts)
        }
        ("narrow", "year") => RelativeTimeFormatter::try_new_narrow_year(&icu_loc.into(), opts),
        // Default style "long" or unknown.
        (_, "second") => RelativeTimeFormatter::try_new_long_second(&icu_loc.into(), opts),
        (_, "minute") => RelativeTimeFormatter::try_new_long_minute(&icu_loc.into(), opts),
        (_, "hour") => RelativeTimeFormatter::try_new_long_hour(&icu_loc.into(), opts),
        (_, "day") => RelativeTimeFormatter::try_new_long_day(&icu_loc.into(), opts),
        (_, "week") => RelativeTimeFormatter::try_new_long_week(&icu_loc.into(), opts),
        (_, "month") => RelativeTimeFormatter::try_new_long_month(&icu_loc.into(), opts),
        (_, "quarter") => RelativeTimeFormatter::try_new_long_quarter(&icu_loc.into(), opts),
        (_, "year") => RelativeTimeFormatter::try_new_long_year(&icu_loc.into(), opts),
        _ => return fallback_relative_time(value, unit),
    };

    let formatter = match formatter_result {
        Ok(f) => f,
        Err(_) => return fallback_relative_time(value, unit),
    };

    // FixedDecimal from f64 — round to integer for typical relative-time
    // use; the fractional part is rare in JS code.
    let decimal: FixedDecimal = (value as i64).into();
    formatter.format(decimal).to_string()
}

fn is_valid_relative_time_unit(unit: &str) -> bool {
    matches!(
        unit.trim_end_matches('s'),
        "second" | "minute" | "hour" | "day" | "week" | "month" | "quarter" | "year"
    )
}

fn format_relative_time_parts(
    locale: &str,
    style: &str,
    numeric: &str,
    value: f64,
    unit: &str,
) -> Vec<Value> {
    if numeric == "auto" {
        return vec![part_obj(
            "literal",
            &format_relative_time_real(locale, style, numeric, value, unit),
        )];
    }
    let unit_norm = unit.trim_end_matches('s');
    let abs = value.abs() as i64;
    let number = abs.to_string();
    let unit_text = if abs == 1 {
        format!(" {}", unit_norm)
    } else {
        format!(" {}s", unit_norm)
    };
    if value.is_sign_negative() {
        vec![
            part_obj("integer", &number),
            part_obj("literal", &format!("{unit_text} ago")),
        ]
    } else {
        vec![
            part_obj("literal", "in "),
            part_obj("integer", &number),
            part_obj("unit", &unit_text),
        ]
    }
}

fn fallback_relative_time(value: f64, unit: &str) -> String {
    let abs = value.abs() as i64;
    let stem = unit.trim_end_matches('s');
    let unit_str = if abs == 1 {
        stem.to_string()
    } else {
        format!("{}s", stem)
    };
    if value.is_sign_negative() {
        format!("{} {} ago", abs, unit_str)
    } else {
        format!("in {} {}", abs, unit_str)
    }
}

// ── Intl.Segmenter (ECMA-402 §18) ────────────────────────────────────
//
// Backed by `unicode-segmentation` — UAX #29 grapheme cluster, word,
// and sentence boundaries. Language-agnostic (works for ALL Unicode
// text including emoji, CJK, RTL).

fn register_segmenter(vm: &mut VM) {
    vm.register_host_fn(
        "ecma:intl/segmenter",
        "new",
        Box::new(|_ctx, args| {
            let locale = resolve_locale(args.first());
            let options = resolve_options(args.get(1));
            let ol = options.lock().unwrap();
            let granularity = ol
                .properties
                .get("granularity")
                .and_then(|v| {
                    if let Value::String(s) = v {
                        Some(s.to_string())
                    } else {
                        None
                    }
                })
                .unwrap_or_else(|| "grapheme".into());
            drop(ol);
            make_object(vec![
                ("__type", s_val("Segmenter")),
                ("__proto__", shared_segmenter_prototype()),
                ("locale", s_val(&locale)),
                ("granularity", s_val(&granularity)),
            ])
        }),
    );

    vm.register_host_fn(
        "ecma:intl/segmenter",
        "segment",
        Box::new(|_ctx, args| {
            use unicode_segmentation::UnicodeSegmentation;
            let seg = match args.first() {
                Some(Value::Object(o)) => o.clone(),
                _ => return make_array(vec![]),
            };
            let input = match args.get(1) {
                Some(Value::String(s)) => s.to_string(),
                Some(o) => format!("{}", o),
                None => String::new(),
            };
            let granularity =
                obj_string_prop(&seg, "granularity").unwrap_or_else(|| "grapheme".into());

            let segments: Vec<(String, usize)> = match granularity.as_str() {
                "word" => input
                    .split_word_bound_indices()
                    .map(|(i, s)| (s.to_string(), i))
                    .collect(),
                "sentence" => input
                    .split_sentence_bound_indices()
                    .map(|(i, s)| (s.to_string(), i))
                    .collect(),
                _ => input
                    .grapheme_indices(true)
                    .map(|(i, s)| (s.to_string(), i))
                    .collect(),
            };

            let elems: Vec<Value> = segments
                .into_iter()
                .map(|(seg_str, idx)| {
                    let is_word_like = granularity == "word" && segment_is_word_like(&seg_str);
                    make_object(vec![
                        ("segment", s_val(&seg_str)),
                        ("index", Value::I32(idx as i32)),
                        ("input", s_val(&input)),
                        ("isWordLike", Value::Bool(is_word_like)),
                    ])
                })
                .collect();
            make_array(elems)
        }),
    );

    // %Segments.prototype%.containing(index) — the segment containing the code
    // unit at `index`, or `undefined` when out of range. The Segments object
    // this operates on is the array `segment` returned, so the lookup is over
    // its elements' `index` + `segment` length.
    vm.register_host_fn(
        "ecma:intl/segmenter",
        "containing",
        Box::new(|_ctx, args| {
            let segments = match args.first() {
                Some(Value::Object(o)) => o.clone(),
                _ => return Value::Undefined,
            };
            let index = match args.get(1) {
                Some(v) => v.as_f64(),
                None => 0.0,
            };
            if !index.is_finite() || index < 0.0 {
                return Value::Undefined;
            }
            let target = index as usize;
            let guard = match segments.lock() {
                Ok(g) => g,
                Err(_) => return Value::Undefined,
            };
            let ObjectKind::Array(elems) = &guard.kind else {
                return Value::Undefined;
            };
            for element in elems.iter() {
                let Value::Object(part) = element else {
                    continue;
                };
                let (start, len) = {
                    let p = match part.lock() {
                        Ok(p) => p,
                        Err(_) => continue,
                    };
                    let start = p
                        .properties
                        .get("index")
                        .map(|v| v.as_f64() as usize)
                        .unwrap_or(0);
                    let len = match p.properties.get("segment") {
                        Some(Value::String(text)) => text.len(),
                        _ => 0,
                    };
                    (start, len)
                };
                if target >= start && target < start + len {
                    return element.clone();
                }
            }
            Value::Undefined
        }),
    );

    vm.register_host_fn(
        "ecma:intl/segmenter",
        "resolvedOptions",
        Box::new(|_ctx, args| {
            if let Some(Value::Object(seg)) = args.first() {
                let locale = obj_string_prop(seg, "locale").unwrap_or_else(|| "en-US".into());
                let granularity =
                    obj_string_prop(seg, "granularity").unwrap_or_else(|| "grapheme".into());
                return make_object(vec![
                    ("locale", s_val(&locale)),
                    ("granularity", s_val(&granularity)),
                ]);
            }
            make_object(vec![])
        }),
    );

    vm.register_host_fn(
        "ecma:intl/segmenter",
        "supportedLocalesOf",
        Box::new(|_ctx, args| match args.first() {
            Some(v @ Value::Object(_)) => v.clone(),
            Some(Value::String(s)) => make_array(vec![Value::String(s.clone())]),
            _ => make_array(vec![]),
        }),
    );
}

fn format_number_parts_real(nf: &Arc<Mutex<Object>>, value: f64) -> Vec<Value> {
    let formatted = format_number_real(nf, value);
    let mut parts = Vec::new();
    let mut seen_decimal = false;
    let chars: Vec<char> = formatted.chars().collect();
    let mut index = 0;

    while index < chars.len() {
        let ch = chars[index];
        if ch.is_ascii_digit() {
            let start = index;
            while index < chars.len() && chars[index].is_ascii_digit() {
                index += 1;
            }
            let value: String = chars[start..index].iter().collect();
            parts.push(make_object(vec![
                (
                    "type",
                    s_val(if seen_decimal { "fraction" } else { "integer" }),
                ),
                ("value", s_val(&value)),
            ]));
            continue;
        }

        let part_type = match ch {
            ',' => "group",
            '.' => {
                seen_decimal = true;
                "decimal"
            }
            '%' => "percentSign",
            '$' | '€' | '£' | '¥' | '¤' => "currency",
            _ => "literal",
        };
        parts.push(make_object(vec![
            ("type", s_val(part_type)),
            ("value", s_val(&ch.to_string())),
        ]));
        index += 1;
    }

    if parts.is_empty() {
        parts.push(make_object(vec![
            ("type", s_val("literal")),
            ("value", s_val(&formatted)),
        ]));
    }
    parts
}

fn format_date_parts_real(dtf: &Arc<Mutex<Object>>, ms: f64) -> Vec<Value> {
    let secs = (ms / 1000.0) as i64;
    let (year, month, day) = epoch_to_ymd(secs);
    let year_opt = obj_string_prop(dtf, "year").unwrap_or_default();
    let month_opt = obj_string_prop(dtf, "month").unwrap_or_default();
    let day_opt = obj_string_prop(dtf, "day").unwrap_or_default();

    if year_opt == "numeric" && month_opt.is_empty() && day_opt.is_empty() {
        return vec![make_object(vec![
            ("type", s_val("year")),
            ("value", s_val(&year.to_string())),
        ])];
    }

    if (year_opt == "numeric" || year_opt == "2-digit")
        && (month_opt == "numeric" || month_opt == "2-digit")
        && (day_opt == "numeric" || day_opt == "2-digit")
    {
        let year_part = if year_opt == "2-digit" {
            format!("{:02}", year.rem_euclid(100))
        } else {
            year.to_string()
        };
        return vec![
            part_obj("month", &format_month_part(month, &month_opt)),
            part_obj("literal", "/"),
            part_obj("day", &format_day_part(day, &day_opt)),
            part_obj("literal", "/"),
            part_obj("year", &year_part),
        ];
    }

    let mut parts = Vec::new();
    if !month_opt.is_empty() {
        parts.push(part_obj("month", &format_month_part(month, &month_opt)));
    }
    if !day_opt.is_empty() {
        if !parts.is_empty() {
            parts.push(part_obj("literal", " "));
        }
        parts.push(part_obj("day", &format_day_part(day, &day_opt)));
    }
    if !year_opt.is_empty() {
        if !parts.is_empty() {
            parts.push(part_obj("literal", ", "));
        }
        let year_part = if year_opt == "2-digit" {
            format!("{:02}", year.rem_euclid(100))
        } else {
            year.to_string()
        };
        parts.push(part_obj("year", &year_part));
    }

    if parts.is_empty() {
        let formatted = format_date_real(dtf, ms);
        if formatted.contains('/') {
            return vec![
                part_obj("month", &month.to_string()),
                part_obj("literal", "/"),
                part_obj("day", &day.to_string()),
                part_obj("literal", "/"),
                part_obj("year", &year.to_string()),
            ];
        }
        parts.push(part_obj("literal", &formatted));
    }
    parts
}

fn format_month_part(month: i32, month_opt: &str) -> String {
    match month_opt {
        "2-digit" => format!("{:02}", month),
        "numeric" => month.to_string(),
        "short" => month_name(month, true).to_string(),
        _ => month_name(month, false).to_string(),
    }
}

fn format_day_part(day: i32, day_opt: &str) -> String {
    if day_opt == "2-digit" {
        format!("{:02}", day)
    } else {
        day.to_string()
    }
}

fn month_name(month: i32, short: bool) -> &'static str {
    match (month, short) {
        (1, false) => "January",
        (2, false) => "February",
        (3, false) => "March",
        (4, false) => "April",
        (5, false) => "May",
        (6, false) => "June",
        (7, false) => "July",
        (8, false) => "August",
        (9, false) => "September",
        (10, false) => "October",
        (11, false) => "November",
        (12, false) => "December",
        (1, true) => "Jan",
        (2, true) => "Feb",
        (3, true) => "Mar",
        (4, true) => "Apr",
        (5, true) => "May",
        (6, true) => "Jun",
        (7, true) => "Jul",
        (8, true) => "Aug",
        (9, true) => "Sep",
        (10, true) => "Oct",
        (11, true) => "Nov",
        (12, true) => "Dec",
        _ => "",
    }
}

fn auto_relative_time_phrase(value: i64, unit: &str) -> Option<String> {
    match (value, unit) {
        (0, "second") => Some("now".into()),
        (-1, "day") => Some("yesterday".into()),
        (0, "day") => Some("today".into()),
        (1, "day") => Some("tomorrow".into()),
        (-1, "week") => Some("last week".into()),
        (0, "week") => Some("this week".into()),
        (1, "week") => Some("next week".into()),
        (-1, "month") => Some("last month".into()),
        (0, "month") => Some("this month".into()),
        (1, "month") => Some("next month".into()),
        (-1, "quarter") => Some("last quarter".into()),
        (0, "quarter") => Some("this quarter".into()),
        (1, "quarter") => Some("next quarter".into()),
        (-1, "year") => Some("last year".into()),
        (0, "year") => Some("this year".into()),
        (1, "year") => Some("next year".into()),
        _ => None,
    }
}

fn segment_is_word_like(segment: &str) -> bool {
    segment.chars().any(|ch| ch.is_alphanumeric())
}

// ── Intl.Locale (ECMA-402 §14) ───────────────────────────────────────
//
// Backed by `unic-langid` for parsing — produces canonical BCP-47
// representation with proper subtag normalization.

fn register_locale(vm: &mut VM) {
    vm.register_host_fn(
        "ecma:intl/locale",
        "new",
        Box::new(|_ctx, args| {
            let tag = match args.first() {
                Some(Value::String(s)) => s.to_string(),
                Some(o) => format!("{}", o),
                None => "en".into(),
            };
            let langid = parse_langid(&tag);
            let language = langid.language.as_str().to_string();
            let region = langid
                .region
                .map(|r| r.as_str().to_string())
                .unwrap_or_default();
            let script = langid
                .script
                .map(|s| s.as_str().to_string())
                .unwrap_or_default();
            let base_name = langid.to_string();
            // Unicode `-u-` extension keywords (UTS #35) and the options bag.
            // These were all hardcoded to "", so `new Intl.Locale(
            // "en-US-u-ca-buddhist-hc-h12")` reported no calendar and no hour
            // cycle. The options bag wins over the tag, per §14.1.2.
            let ext = unicode_extension_keywords(&tag);
            let opt = |key: &str| -> Option<String> {
                args.get(1).and_then(|v| match v {
                    Value::Object(o) => obj_string_prop(o, key),
                    _ => None,
                })
            };
            let keyword = |kw: &str, opt_key: &str| -> String {
                opt(opt_key)
                    .or_else(|| ext.get(kw).cloned())
                    .unwrap_or_default()
            };
            let numeric = opt("numeric").map(|v| v == "true").unwrap_or_else(|| {
                ext.get("kn").map(|v| v.is_empty() || v == "true") == Some(true)
            });
            make_object(vec![
                ("__type", s_val("Locale")),
                ("baseName", s_val(&base_name)),
                ("language", s_val(&language)),
                ("region", s_val(&region)),
                ("script", s_val(&script)),
                ("calendar", s_val(&keyword("ca", "calendar"))),
                ("numberingSystem", s_val(&keyword("nu", "numberingSystem"))),
                ("collation", s_val(&keyword("co", "collation"))),
                ("caseFirst", s_val(&keyword("kf", "caseFirst"))),
                ("hourCycle", s_val(&keyword("hc", "hourCycle"))),
                ("numeric", Value::Bool(numeric)),
                // ES2024 §14.3.x — `-u-fw-<day>` and the `firstDayOfWeek`
                // option. Normalized to the spec's Mon=1 … Sun=7 numbering,
                // accepting both the CLDR day codes and a numeric value.
                (
                    "firstDayOfWeek",
                    match keyword("fw", "firstDayOfWeek").as_str() {
                        "" => Value::Undefined,
                        "mon" => Value::F64(1.0),
                        "tue" => Value::F64(2.0),
                        "wed" => Value::F64(3.0),
                        "thu" => Value::F64(4.0),
                        "fri" => Value::F64(5.0),
                        "sat" => Value::F64(6.0),
                        "sun" => Value::F64(7.0),
                        other => match other.parse::<f64>() {
                            Ok(n) if (1.0..=7.0).contains(&n) => Value::F64(n),
                            _ => Value::Undefined,
                        },
                    },
                ),
                // BCP-47 variant subtags, in canonical (lower-case, sorted)
                // form; absent when the tag carries none.
                ("variants", {
                    let variants: Vec<String> =
                        langid.variants().map(|v| v.as_str().to_string()).collect();
                    if variants.is_empty() {
                        Value::Undefined
                    } else {
                        s_val(&variants.join("-"))
                    }
                }),
            ])
        }),
    );

    // ── ES2024 Locale Info (§14.3.x) ──────────────────────────────────────
    // Each getter returns the locale's resolved preference FIRST when the tag
    // pinned one via `-u-`, then the locale-independent default. `getTimeZones`
    // is the one backed by real data now that tzdb is linked.
    for (fn_name, keyword, defaults) in [
        ("getCalendars", "calendar", &["gregory"][..]),
        ("getCollations", "collation", &["default"][..]),
        ("getNumberingSystems", "numberingSystem", &["latn"][..]),
        ("getHourCycles", "hourCycle", &["h23"][..]),
    ] {
        let keyword = keyword.to_string();
        let defaults: Vec<String> = defaults.iter().map(|s| s.to_string()).collect();
        vm.register_host_fn(
            "ecma:intl/locale",
            fn_name,
            Box::new(move |_ctx, args| {
                if let Some(Value::Object(loc)) = args.first() {
                    if let Some(resolved) = obj_string_prop(loc, &keyword) {
                        if !resolved.is_empty() {
                            return make_array(vec![s_val(&resolved)]);
                        }
                    }
                }
                make_array(defaults.iter().map(|d| s_val(d)).collect())
            }),
        );
    }

    // getTimeZones() → the IANA identifiers for the locale's region, or
    // `undefined` when the locale carries no region, per §14.3.x. Backed by the
    // vendored tzdb `zone.tab` (public domain) — the mapping exists in neither
    // `chrono-tz` nor ICU.
    vm.register_host_fn(
        "ecma:intl/locale",
        "getTimeZones",
        Box::new(|_ctx, args| {
            let region = match args.first() {
                Some(Value::Object(loc)) => obj_string_prop(loc, "region").unwrap_or_default(),
                _ => String::new(),
            };
            if region.is_empty() {
                return Value::Undefined;
            }
            make_array(
                crate::timezone::identifiers_for_region(&region)
                    .into_iter()
                    .map(|zone| s_val(&zone))
                    .collect(),
            )
        }),
    );

    // getTextInfo() → { direction }, from CLDR script-direction data.
    vm.register_host_fn(
        "ecma:intl/locale",
        "getTextInfo",
        Box::new(|_ctx, args| {
            use icu::locale::{Direction, LanguageIdentifier, LocaleDirectionality};
            let tag = match args.first() {
                Some(Value::Object(loc)) => obj_string_prop(loc, "baseName").unwrap_or_default(),
                _ => String::new(),
            };
            // NOTE: `parse_langid` here returns `unic_langid`'s type, which is a
            // DIFFERENT crate from ICU's — parse with ICU's own.
            let langid: LanguageIdentifier = tag.parse().unwrap_or(LanguageIdentifier::UNKNOWN);
            let direction = match LocaleDirectionality::new_common().get(&langid) {
                Some(Direction::RightToLeft) => "rtl",
                _ => "ltr",
            };
            make_object(vec![("direction", s_val(direction))])
        }),
    );

    // getWeekInfo() → { firstDay, weekend, minimalDays }, from CLDR week data.
    // `firstDay`/`weekend` use ECMA-402's numbering (Mon=1 … Sun=7).
    vm.register_host_fn(
        "ecma:intl/locale",
        "getWeekInfo",
        Box::new(|_ctx, args| {
            use icu::calendar::week::WeekInformation;
            use icu::locale::Locale;
            let tag = match args.first() {
                Some(Value::Object(loc)) => obj_string_prop(loc, "baseName").unwrap_or_default(),
                _ => String::new(),
            };
            let locale: Locale = tag.parse().unwrap_or_else(|_| Locale::UNKNOWN);
            let Ok(info) = WeekInformation::try_new(locale.into()) else {
                return Value::Undefined;
            };
            // `Weekday` is Mon=1-based already, matching ECMA-402.
            let day_number = |d: icu::calendar::types::Weekday| d as u8 as f64;
            make_object(vec![
                ("firstDay", Value::F64(day_number(info.first_weekday))),
                (
                    "weekend",
                    make_array(info.weekend().map(|d| Value::F64(day_number(d))).collect()),
                ),
                // CLDR's minDays is not exposed on WeekInformation; ISO 8601's
                // value is the ECMA-402 default and what ICU's own calculator
                // uses (`WeekCalculator::ISO`).
                ("minimalDays", Value::F64(4.0)),
            ])
        }),
    );

    vm.register_host_fn(
        "ecma:intl/locale",
        "toString",
        Box::new(|_ctx, args| {
            if let Some(Value::Object(loc)) = args.first() {
                return s_val(&obj_string_prop(loc, "baseName").unwrap_or_default());
            }
            s_val("")
        }),
    );

    vm.register_host_fn(
        "ecma:intl/locale",
        "maximize",
        Box::new(|_ctx, args| {
            // Likely-subtag expansion needs CLDR data; unic-langid 0.9
            // alone doesn't provide it. Return as-is — full expansion is
            // a future enhancement (could use icu::locale::Locale::maximize).
            args.first().cloned().unwrap_or(Value::Null)
        }),
    );

    vm.register_host_fn(
        "ecma:intl/locale",
        "minimize",
        Box::new(|_ctx, args| args.first().cloned().unwrap_or(Value::Null)),
    );
}

// ── Intl.DisplayNames (ECMA-402 §12) ─────────────────────────────────
//
// Backed by `icu_displaynames` — full CLDR translation tables.

fn register_display_names(vm: &mut VM) {
    vm.register_host_fn(
        "ecma:intl/displaynames",
        "new",
        Box::new(|_ctx, args| {
            let locale = resolve_locale(args.first());
            let options = resolve_options(args.get(1));
            let ol = options.lock().unwrap();
            let dn_type = ol
                .properties
                .get("type")
                .and_then(|v| {
                    if let Value::String(s) = v {
                        Some(s.to_string())
                    } else {
                        None
                    }
                })
                .unwrap_or_else(|| "language".into());
            let style = ol
                .properties
                .get("style")
                .and_then(|v| {
                    if let Value::String(s) = v {
                        Some(s.to_string())
                    } else {
                        None
                    }
                })
                .unwrap_or_else(|| "long".into());
            drop(ol);
            make_object(vec![
                ("__type", s_val("DisplayNames")),
                ("locale", s_val(&locale)),
                ("type", s_val(&dn_type)),
                ("style", s_val(&style)),
            ])
        }),
    );

    vm.register_host_fn(
        "ecma:intl/displaynames",
        "of",
        Box::new(|_ctx, args| {
            let dn = match args.first() {
                Some(Value::Object(o)) => o.clone(),
                _ => return Value::Undefined,
            };
            let code = match args.get(1) {
                Some(Value::String(s)) => s.to_string(),
                Some(o) => format!("{}", o),
                None => return Value::Undefined,
            };
            let locale = obj_string_prop(&dn, "locale").unwrap_or_else(|| "en-US".into());
            let dn_type = obj_string_prop(&dn, "type").unwrap_or_else(|| "language".into());
            s_val(&display_name_of(&locale, &dn_type, &code))
        }),
    );

    vm.register_host_fn(
        "ecma:intl/displaynames",
        "resolvedOptions",
        Box::new(|_ctx, args| {
            if let Some(Value::Object(dn)) = args.first() {
                let locale = obj_string_prop(dn, "locale").unwrap_or_else(|| "en-US".into());
                let dn_type = obj_string_prop(dn, "type").unwrap_or_else(|| "language".into());
                let style = obj_string_prop(dn, "style").unwrap_or_else(|| "long".into());
                return make_object(vec![
                    ("locale", s_val(&locale)),
                    ("type", s_val(&dn_type)),
                    ("style", s_val(&style)),
                    ("fallback", s_val("code")),
                ]);
            }
            make_object(vec![])
        }),
    );

    vm.register_host_fn(
        "ecma:intl/displaynames",
        "supportedLocalesOf",
        Box::new(|_ctx, args| match args.first() {
            Some(v @ Value::Object(_)) => v.clone(),
            Some(Value::String(s)) => make_array(vec![Value::String(s.clone())]),
            _ => make_array(vec![]),
        }),
    );
}

/// Return the display name for `code` in the given locale's perspective.
/// `dn_type` is "language" / "region" / "script" / "currency".
///
/// `icu_displaynames` 0.11 has the API but its locale parsing API uses
/// the older `icu_locid` (transitive dep). We bridge by parsing the
/// tag through that crate.
fn display_name_of(locale: &str, dn_type: &str, code: &str) -> String {
    use icu_displaynames::{
        DisplayNamesOptions, LanguageDisplayNames, RegionDisplayNames, ScriptDisplayNames,
    };

    let icu_loc: icu_locid::Locale = match locale.parse() {
        Ok(l) => l,
        Err(_) => return code.to_string(),
    };

    let opts = DisplayNamesOptions::default();
    match dn_type {
        "language" => {
            if let Ok(dn) = LanguageDisplayNames::try_new(&icu_loc.into(), opts) {
                if let Ok(lang_id) = code.parse::<icu_locid::subtags::Language>() {
                    if let Some(name) = dn.of(lang_id) {
                        return name.to_string();
                    }
                }
            }
            code.to_string()
        }
        "region" => {
            if let Ok(dn) = RegionDisplayNames::try_new(&icu_loc.into(), opts) {
                if let Ok(region) = code.to_uppercase().parse::<icu_locid::subtags::Region>() {
                    if let Some(name) = dn.of(region) {
                        return name.to_string();
                    }
                }
            }
            code.to_string()
        }
        "script" => {
            if let Ok(dn) = ScriptDisplayNames::try_new(&icu_loc.into(), opts) {
                if let Ok(script) = code.parse::<icu_locid::subtags::Script>() {
                    if let Some(name) = dn.of(script) {
                        return name.to_string();
                    }
                }
            }
            code.to_string()
        }
        _ => code.to_string(),
    }
}

// ── Intl.DurationFormat (ECMA-402 §19, Stage 4 in 2024) ──────────────
//
// `icu_experimental::duration` is still pre-1.0 in icu4x. MVP keeps the
// hand-rolled English formatter until that crate stabilises.

fn register_duration_format(vm: &mut VM) {
    vm.register_host_fn(
        "ecma:intl/durationformat",
        "new",
        Box::new(|_ctx, args| {
            let locale = resolve_locale(args.first());
            let options = resolve_options(args.get(1));
            let ol = options.lock().unwrap();
            let style = ol
                .properties
                .get("style")
                .and_then(|v| {
                    if let Value::String(s) = v {
                        Some(s.to_string())
                    } else {
                        None
                    }
                })
                .unwrap_or_else(|| "short".into());
            drop(ol);
            make_object(vec![
                ("__type", s_val("DurationFormat")),
                ("locale", s_val(&locale)),
                ("style", s_val(&style)),
            ])
        }),
    );

    vm.register_host_fn(
        "ecma:intl/durationformat",
        "format",
        Box::new(|_ctx, args| {
            let df = match args.first() {
                Some(Value::Object(o)) => o.clone(),
                _ => return s_val(""),
            };
            let dur = match args.get(1) {
                Some(Value::Object(o)) => o.clone(),
                _ => return s_val(""),
            };
            let style = obj_string_prop(&df, "style").unwrap_or_else(|| "short".into());
            let locale = obj_string_prop(&df, "locale").unwrap_or_else(|| "en-US".into());
            s_val(&format_duration(&dur, &style, &locale))
        }),
    );

    vm.register_host_fn(
        "ecma:intl/durationformat",
        "formatToParts",
        Box::new(|_ctx, args| {
            let df = match args.first() {
                Some(Value::Object(o)) => o.clone(),
                _ => return make_array(vec![]),
            };
            let dur = match args.get(1) {
                Some(Value::Object(o)) => o.clone(),
                _ => return make_array(vec![]),
            };
            let style = obj_string_prop(&df, "style").unwrap_or_else(|| "short".into());
            let locale = obj_string_prop(&df, "locale").unwrap_or_else(|| "en-US".into());
            make_array(vec![make_object(vec![
                ("type", s_val("literal")),
                ("value", s_val(&format_duration(&dur, &style, &locale))),
            ])])
        }),
    );

    vm.register_host_fn(
        "ecma:intl/durationformat",
        "resolvedOptions",
        Box::new(|_ctx, args| {
            if let Some(Value::Object(df)) = args.first() {
                let locale = obj_string_prop(df, "locale").unwrap_or_else(|| "en-US".into());
                let style = obj_string_prop(df, "style").unwrap_or_else(|| "short".into());
                return make_object(vec![
                    ("locale", s_val(&locale)),
                    ("style", s_val(&style)),
                    ("numberingSystem", s_val("latn")),
                ]);
            }
            make_object(vec![])
        }),
    );

    vm.register_host_fn(
        "ecma:intl/durationformat",
        "supportedLocalesOf",
        Box::new(|_ctx, args| match args.first() {
            Some(v @ Value::Object(_)) => v.clone(),
            Some(Value::String(s)) => make_array(vec![Value::String(s.clone())]),
            _ => make_array(vec![]),
        }),
    );
}

/// Format a duration object per ECMA-402 §19 Intl.DurationFormat.
///
/// Supports all 4 spec styles (long / short / narrow / digital) and
/// all 10 spec units (years through nanoseconds). Locale-aware text
/// for the top-20 languages by usage; falls back to English for
/// unknown locales (matches ECMA-402's "best fit" behaviour).
///
/// `digital` style renders hours/minutes/seconds as `H:MM:SS` per the
/// spec's clock-time convention; other units remain in short form.
fn format_duration(dur: &Arc<Mutex<Object>>, style: &str, locale: &str) -> String {
    use intl_pluralrules::{PluralCategory, PluralRuleType, PluralRules};

    // Plural categorizer for this locale — used by long-style
    // localized unit names (Russian few/many, Welsh two, Arabic 6
    // categories, etc. all need the right form).
    let parsed = parse_langid(locale);
    let langid = unic_langid::LanguageIdentifier::from_parts(parsed.language, None, None, &[]);
    let lang_code = parsed.language.as_str();
    let plural_for = |n: i64| -> PluralCategory {
        match PluralRules::create(langid.clone(), PluralRuleType::CARDINAL) {
            Ok(pr) => pr.select(n as isize).unwrap_or(PluralCategory::OTHER),
            Err(_) => PluralCategory::OTHER,
        }
    };

    let lock = dur.lock().unwrap();

    // Digital style: H:MM:SS for hours/minutes/seconds, then short
    // text-form for any other non-zero units.
    if style == "digital" {
        let h = lock
            .properties
            .get("hours")
            .map(|v| v.as_i32())
            .unwrap_or(0);
        let m = lock
            .properties
            .get("minutes")
            .map(|v| v.as_i32())
            .unwrap_or(0);
        let s = lock
            .properties
            .get("seconds")
            .map(|v| v.as_i32())
            .unwrap_or(0);
        let mut clock = format!("{}:{:02}:{:02}", h, m.abs(), s.abs());
        let mut extras = Vec::new();
        for key in &[
            "years",
            "months",
            "weeks",
            "days",
            "milliseconds",
            "microseconds",
            "nanoseconds",
        ] {
            if let Some(v) = lock.properties.get(*key) {
                let n = v.as_i32() as i64;
                if n != 0 {
                    let unit_text = duration_unit_text(lang_code, key, "short", plural_for(n));
                    extras.push(format!("{} {}", n, unit_text));
                }
            }
        }
        if !extras.is_empty() {
            clock = format!("{} {}", extras.join(", "), clock);
        }
        return clock;
    }

    // Long / short / narrow — concat all non-zero units with the
    // localized form for the count's plural category.
    let mut parts = Vec::new();
    for key in &[
        "years",
        "months",
        "weeks",
        "days",
        "hours",
        "minutes",
        "seconds",
        "milliseconds",
        "microseconds",
        "nanoseconds",
    ] {
        if let Some(v) = lock.properties.get(*key) {
            let n = v.as_i32() as i64;
            if n != 0 {
                let unit_text = duration_unit_text(lang_code, key, style, plural_for(n));
                parts.push(format!("{} {}", n, unit_text));
            }
        }
    }
    parts.join(", ")
}

/// Look up the localized unit name for a duration component.
///
/// Table covers the top-20 languages by speaker count + web usage.
/// Each entry is `(lang, unit, style) -> [one_form, other_form]`.
/// Languages with more plural categories (Russian few/many, Arabic
/// zero/one/two/few/many/other) get extra entries.
///
/// Falls back to English if the locale isn't in the table (ECMA-402
/// "best fit" behaviour).
fn duration_unit_text(
    lang: &str,
    unit: &str,
    style: &str,
    category: intl_pluralrules::PluralCategory,
) -> &'static str {
    use intl_pluralrules::PluralCategory::*;
    // Table layout: each (lang, unit, style) entry is an array of forms
    // ordered by category index. We use a closure for lookup so the
    // match is contiguous and the table stays readable.
    let cat_idx = match category {
        ZERO => 0,
        ONE => 1,
        TWO => 2,
        FEW => 3,
        MANY => 4,
        OTHER => 5,
    };
    let forms = duration_forms(lang, unit, style)
        .or_else(|| duration_forms("en", unit, style))
        .unwrap_or(&["", "", "", "", "", ""]);
    // Pick the requested category, falling back to "other" (idx 5)
    // if the language doesn't have that specific form, then to "one".
    if cat_idx < forms.len() && !forms[cat_idx].is_empty() {
        forms[cat_idx]
    } else if !forms[5].is_empty() {
        forms[5]
    } else {
        forms[1]
    }
}

/// Per-(lang, unit, style) form table.
///
/// Forms are ordered: [zero, one, two, few, many, other]. Empty
/// strings mean "not specified — fall through to other".
///
/// Sources:
/// - English/German/French/Spanish/Italian/Portuguese: standard CLDR forms.
/// - Russian/Polish: 3-way plural (one/few/many) per CLDR.
/// - Arabic: 6-category plural per CLDR.
/// - CJK (Chinese/Japanese/Korean): no plural inflection.
/// - Others: standard CLDR singular/plural.
///
/// Short and narrow forms are mostly invariant per language (no plural).
fn duration_forms(lang: &str, unit: &str, style: &str) -> Option<&'static [&'static str; 6]> {
    Some(match (lang, unit, style) {
        // ─── English (en) ─────────────────────────────────────────
        ("en", "years", "long") => &["", "year", "", "", "", "years"],
        ("en", "years", "short") => &["", "yr", "", "", "", "yrs"],
        ("en", "years", "narrow") => &["", "y", "", "", "", "y"],
        ("en", "months", "long") => &["", "month", "", "", "", "months"],
        ("en", "months", "short") => &["", "mo", "", "", "", "mo"],
        ("en", "months", "narrow") => &["", "mo", "", "", "", "mo"],
        ("en", "weeks", "long") => &["", "week", "", "", "", "weeks"],
        ("en", "weeks", "short") => &["", "wk", "", "", "", "wks"],
        ("en", "weeks", "narrow") => &["", "w", "", "", "", "w"],
        ("en", "days", "long") => &["", "day", "", "", "", "days"],
        ("en", "days", "short") => &["", "day", "", "", "", "days"],
        ("en", "days", "narrow") => &["", "d", "", "", "", "d"],
        ("en", "hours", "long") => &["", "hour", "", "", "", "hours"],
        ("en", "hours", "short") => &["", "hr", "", "", "", "hr"],
        ("en", "hours", "narrow") => &["", "h", "", "", "", "h"],
        ("en", "minutes", "long") => &["", "minute", "", "", "", "minutes"],
        ("en", "minutes", "short") => &["", "min", "", "", "", "min"],
        ("en", "minutes", "narrow") => &["", "m", "", "", "", "m"],
        ("en", "seconds", "long") => &["", "second", "", "", "", "seconds"],
        ("en", "seconds", "short") => &["", "sec", "", "", "", "sec"],
        ("en", "seconds", "narrow") => &["", "s", "", "", "", "s"],
        ("en", "milliseconds", "long") => &["", "millisecond", "", "", "", "milliseconds"],
        ("en", "milliseconds", "short") => &["", "ms", "", "", "", "ms"],
        ("en", "milliseconds", "narrow") => &["", "ms", "", "", "", "ms"],
        ("en", "microseconds", "long") => &["", "microsecond", "", "", "", "microseconds"],
        ("en", "microseconds", "short") => &["", "μs", "", "", "", "μs"],
        ("en", "microseconds", "narrow") => &["", "μs", "", "", "", "μs"],
        ("en", "nanoseconds", "long") => &["", "nanosecond", "", "", "", "nanoseconds"],
        ("en", "nanoseconds", "short") => &["", "ns", "", "", "", "ns"],
        ("en", "nanoseconds", "narrow") => &["", "ns", "", "", "", "ns"],

        // ─── Mandarin Chinese (zh) — no plural inflection ─────────
        ("zh", "years", _) => &["", "年", "", "", "", "年"],
        ("zh", "months", _) => &["", "个月", "", "", "", "个月"],
        ("zh", "weeks", _) => &["", "周", "", "", "", "周"],
        ("zh", "days", _) => &["", "天", "", "", "", "天"],
        ("zh", "hours", _) => &["", "小时", "", "", "", "小时"],
        ("zh", "minutes", _) => &["", "分钟", "", "", "", "分钟"],
        ("zh", "seconds", _) => &["", "秒", "", "", "", "秒"],
        ("zh", "milliseconds", _) => &["", "毫秒", "", "", "", "毫秒"],
        ("zh", "microseconds", _) => &["", "微秒", "", "", "", "微秒"],
        ("zh", "nanoseconds", _) => &["", "纳秒", "", "", "", "纳秒"],

        // ─── Spanish (es) ─────────────────────────────────────────
        ("es", "years", "long") => &["", "año", "", "", "", "años"],
        ("es", "years", "short") => &["", "a", "", "", "", "a"],
        ("es", "years", "narrow") => &["", "a", "", "", "", "a"],
        ("es", "months", "long") => &["", "mes", "", "", "", "meses"],
        ("es", "months", _) => &["", "m", "", "", "", "m"],
        ("es", "weeks", "long") => &["", "semana", "", "", "", "semanas"],
        ("es", "weeks", _) => &["", "sem.", "", "", "", "sem."],
        ("es", "days", "long") => &["", "día", "", "", "", "días"],
        ("es", "days", _) => &["", "d", "", "", "", "d"],
        ("es", "hours", "long") => &["", "hora", "", "", "", "horas"],
        ("es", "hours", _) => &["", "h", "", "", "", "h"],
        ("es", "minutes", "long") => &["", "minuto", "", "", "", "minutos"],
        ("es", "minutes", _) => &["", "min", "", "", "", "min"],
        ("es", "seconds", "long") => &["", "segundo", "", "", "", "segundos"],
        ("es", "seconds", _) => &["", "s", "", "", "", "s"],
        ("es", "milliseconds", "long") => &["", "milisegundo", "", "", "", "milisegundos"],
        ("es", "milliseconds", _) => &["", "ms", "", "", "", "ms"],
        ("es", "microseconds", "long") => &["", "microsegundo", "", "", "", "microsegundos"],
        ("es", "microseconds", _) => &["", "μs", "", "", "", "μs"],
        ("es", "nanoseconds", "long") => &["", "nanosegundo", "", "", "", "nanosegundos"],
        ("es", "nanoseconds", _) => &["", "ns", "", "", "", "ns"],

        // ─── Hindi (hi) ───────────────────────────────────────────
        ("hi", "years", "long") => &["", "वर्ष", "", "", "", "वर्ष"],
        ("hi", "months", "long") => &["", "महीना", "", "", "", "महीने"],
        ("hi", "weeks", "long") => &["", "सप्ताह", "", "", "", "सप्ताह"],
        ("hi", "days", "long") => &["", "दिन", "", "", "", "दिन"],
        ("hi", "hours", "long") => &["", "घंटा", "", "", "", "घंटे"],
        ("hi", "minutes", "long") => &["", "मिनट", "", "", "", "मिनट"],
        ("hi", "seconds", "long") => &["", "सेकंड", "", "", "", "सेकंड"],
        ("hi", "milliseconds", "long") => &["", "मिलीसेकंड", "", "", "", "मिलीसेकंड"],
        ("hi", "microseconds", "long") => &["", "माइक्रोसेकंड", "", "", "", "माइक्रोसेकंड"],
        ("hi", "nanoseconds", "long") => &["", "नैनोसेकंड", "", "", "", "नैनोसेकंड"],

        // ─── Arabic (ar) — 6 plural categories ────────────────────
        // CLDR forms: zero, one, two, few (3-10), many (11+), other.
        // Modern Standard Arabic uses different word forms per category.
        ("ar", "years", "long") => &[
            "سنة",   // zero (لا توجد سنوات)
            "سنة",   // one (سنة واحدة)
            "سنتان", // two (سنتان)
            "سنوات", // few (3-10 sanawat)
            "سنة",   // many (11+ "sana")
            "سنة",   // other (decimals)
        ],
        ("ar", "years", _) => &["", "سنة", "", "", "", "سنة"],
        ("ar", "months", "long") => &["شهر", "شهر", "شهران", "أشهر", "شهرًا", "شهر"],
        ("ar", "months", _) => &["", "شهر", "", "", "", "شهر"],
        ("ar", "weeks", "long") => &["أسبوع", "أسبوع", "أسبوعان", "أسابيع", "أسبوعًا", "أسبوع"],
        ("ar", "weeks", _) => &["", "أسبوع", "", "", "", "أسبوع"],
        ("ar", "days", "long") => &["يوم", "يوم", "يومان", "أيام", "يومًا", "يوم"],
        ("ar", "days", _) => &["", "يوم", "", "", "", "يوم"],
        ("ar", "hours", "long") => &["ساعة", "ساعة", "ساعتان", "ساعات", "ساعة", "ساعة"],
        ("ar", "hours", _) => &["", "س", "", "", "", "س"],
        ("ar", "minutes", "long") => &["دقيقة", "دقيقة", "دقيقتان", "دقائق", "دقيقة", "دقيقة"],
        ("ar", "minutes", _) => &["", "د", "", "", "", "د"],
        ("ar", "seconds", "long") => &["ثانية", "ثانية", "ثانيتان", "ثوانٍ", "ثانية", "ثانية"],
        ("ar", "seconds", _) => &["", "ث", "", "", "", "ث"],
        ("ar", "milliseconds", "long") => &["", "ميلي ثانية", "", "", "", "ميلي ثانية"],
        ("ar", "milliseconds", _) => &["", "ms", "", "", "", "ms"],
        ("ar", "microseconds", "long") => &["", "ميكروثانية", "", "", "", "ميكروثانية"],
        ("ar", "microseconds", _) => &["", "μs", "", "", "", "μs"],
        ("ar", "nanoseconds", "long") => &["", "نانوثانية", "", "", "", "نانوثانية"],
        ("ar", "nanoseconds", _) => &["", "ns", "", "", "", "ns"],

        // ─── Portuguese (pt) ──────────────────────────────────────
        ("pt", "years", "long") => &["", "ano", "", "", "", "anos"],
        ("pt", "years", _) => &["", "a", "", "", "", "a"],
        ("pt", "months", "long") => &["", "mês", "", "", "", "meses"],
        ("pt", "months", _) => &["", "m", "", "", "", "m"],
        ("pt", "weeks", "long") => &["", "semana", "", "", "", "semanas"],
        ("pt", "weeks", _) => &["", "sem.", "", "", "", "sem."],
        ("pt", "days", "long") => &["", "dia", "", "", "", "dias"],
        ("pt", "days", _) => &["", "d", "", "", "", "d"],
        ("pt", "hours", "long") => &["", "hora", "", "", "", "horas"],
        ("pt", "hours", _) => &["", "h", "", "", "", "h"],
        ("pt", "minutes", "long") => &["", "minuto", "", "", "", "minutos"],
        ("pt", "minutes", _) => &["", "min", "", "", "", "min"],
        ("pt", "seconds", "long") => &["", "segundo", "", "", "", "segundos"],
        ("pt", "seconds", _) => &["", "s", "", "", "", "s"],
        ("pt", "milliseconds", "long") => &["", "milissegundo", "", "", "", "milissegundos"],
        ("pt", "milliseconds", _) => &["", "ms", "", "", "", "ms"],

        // ─── Bengali (bn) ─────────────────────────────────────────
        ("bn", "years", "long") => &["", "বছর", "", "", "", "বছর"],
        ("bn", "months", "long") => &["", "মাস", "", "", "", "মাস"],
        ("bn", "weeks", "long") => &["", "সপ্তাহ", "", "", "", "সপ্তাহ"],
        ("bn", "days", "long") => &["", "দিন", "", "", "", "দিন"],
        ("bn", "hours", "long") => &["", "ঘণ্টা", "", "", "", "ঘণ্টা"],
        ("bn", "minutes", "long") => &["", "মিনিট", "", "", "", "মিনিট"],
        ("bn", "seconds", "long") => &["", "সেকেন্ড", "", "", "", "সেকেন্ড"],

        // ─── Russian (ru) — 3-way plural ──────────────────────────
        // Forms ordered: [_, one, _, few, many, other]
        // one: 1, 21, 31...   few: 2-4, 22-24...   many: 0, 5-20...
        ("ru", "years", "long") => &["", "год", "", "года", "лет", "года"],
        ("ru", "years", _) => &["", "г.", "", "г.", "г.", "г."],
        ("ru", "months", "long") => &["", "месяц", "", "месяца", "месяцев", "месяца"],
        ("ru", "months", _) => &["", "мес.", "", "мес.", "мес.", "мес."],
        ("ru", "weeks", "long") => &["", "неделя", "", "недели", "недель", "недели"],
        ("ru", "weeks", _) => &["", "нед.", "", "нед.", "нед.", "нед."],
        ("ru", "days", "long") => &["", "день", "", "дня", "дней", "дня"],
        ("ru", "days", _) => &["", "дн.", "", "дн.", "дн.", "дн."],
        ("ru", "hours", "long") => &["", "час", "", "часа", "часов", "часа"],
        ("ru", "hours", _) => &["", "ч.", "", "ч.", "ч.", "ч."],
        ("ru", "minutes", "long") => &["", "минута", "", "минуты", "минут", "минуты"],
        ("ru", "minutes", _) => &["", "мин.", "", "мин.", "мин.", "мин."],
        ("ru", "seconds", "long") => &["", "секунда", "", "секунды", "секунд", "секунды"],
        ("ru", "seconds", _) => &["", "сек.", "", "сек.", "сек.", "сек."],
        ("ru", "milliseconds", "long") => &[
            "",
            "миллисекунда",
            "",
            "миллисекунды",
            "миллисекунд",
            "миллисекунды",
        ],
        ("ru", "milliseconds", _) => &["", "мс", "", "мс", "мс", "мс"],

        // ─── Japanese (ja) — no plural inflection ─────────────────
        ("ja", "years", _) => &["", "年", "", "", "", "年"],
        ("ja", "months", _) => &["", "か月", "", "", "", "か月"],
        ("ja", "weeks", _) => &["", "週間", "", "", "", "週間"],
        ("ja", "days", _) => &["", "日", "", "", "", "日"],
        ("ja", "hours", _) => &["", "時間", "", "", "", "時間"],
        ("ja", "minutes", _) => &["", "分", "", "", "", "分"],
        ("ja", "seconds", _) => &["", "秒", "", "", "", "秒"],
        ("ja", "milliseconds", _) => &["", "ミリ秒", "", "", "", "ミリ秒"],
        ("ja", "microseconds", _) => &["", "マイクロ秒", "", "", "", "マイクロ秒"],
        ("ja", "nanoseconds", _) => &["", "ナノ秒", "", "", "", "ナノ秒"],

        // ─── German (de) ──────────────────────────────────────────
        ("de", "years", "long") => &["", "Jahr", "", "", "", "Jahre"],
        ("de", "years", _) => &["", "J", "", "", "", "J"],
        ("de", "months", "long") => &["", "Monat", "", "", "", "Monate"],
        ("de", "months", _) => &["", "Mon.", "", "", "", "Mon."],
        ("de", "weeks", "long") => &["", "Woche", "", "", "", "Wochen"],
        ("de", "weeks", _) => &["", "Wo.", "", "", "", "Wo."],
        ("de", "days", "long") => &["", "Tag", "", "", "", "Tage"],
        ("de", "days", _) => &["", "Tg.", "", "", "", "Tg."],
        ("de", "hours", "long") => &["", "Stunde", "", "", "", "Stunden"],
        ("de", "hours", _) => &["", "Std.", "", "", "", "Std."],
        ("de", "minutes", "long") => &["", "Minute", "", "", "", "Minuten"],
        ("de", "minutes", _) => &["", "Min.", "", "", "", "Min."],
        ("de", "seconds", "long") => &["", "Sekunde", "", "", "", "Sekunden"],
        ("de", "seconds", _) => &["", "Sek.", "", "", "", "Sek."],
        ("de", "milliseconds", "long") => &["", "Millisekunde", "", "", "", "Millisekunden"],
        ("de", "milliseconds", _) => &["", "ms", "", "", "", "ms"],

        // ─── French (fr) — `one` covers 0 and 1 ───────────────────
        ("fr", "years", "long") => &["", "an", "", "", "", "ans"],
        ("fr", "years", _) => &["", "an", "", "", "", "ans"],
        ("fr", "months", "long") => &["", "mois", "", "", "", "mois"],
        ("fr", "months", _) => &["", "m.", "", "", "", "m."],
        ("fr", "weeks", "long") => &["", "semaine", "", "", "", "semaines"],
        ("fr", "weeks", _) => &["", "sem.", "", "", "", "sem."],
        ("fr", "days", "long") => &["", "jour", "", "", "", "jours"],
        ("fr", "days", _) => &["", "j", "", "", "", "j"],
        ("fr", "hours", "long") => &["", "heure", "", "", "", "heures"],
        ("fr", "hours", _) => &["", "h", "", "", "", "h"],
        ("fr", "minutes", "long") => &["", "minute", "", "", "", "minutes"],
        ("fr", "minutes", _) => &["", "min", "", "", "", "min"],
        ("fr", "seconds", "long") => &["", "seconde", "", "", "", "secondes"],
        ("fr", "seconds", _) => &["", "s", "", "", "", "s"],
        ("fr", "milliseconds", "long") => &["", "milliseconde", "", "", "", "millisecondes"],
        ("fr", "milliseconds", _) => &["", "ms", "", "", "", "ms"],

        // ─── Korean (ko) — no plural inflection ───────────────────
        ("ko", "years", _) => &["", "년", "", "", "", "년"],
        ("ko", "months", _) => &["", "개월", "", "", "", "개월"],
        ("ko", "weeks", _) => &["", "주", "", "", "", "주"],
        ("ko", "days", _) => &["", "일", "", "", "", "일"],
        ("ko", "hours", _) => &["", "시간", "", "", "", "시간"],
        ("ko", "minutes", _) => &["", "분", "", "", "", "분"],
        ("ko", "seconds", _) => &["", "초", "", "", "", "초"],
        ("ko", "milliseconds", _) => &["", "밀리초", "", "", "", "밀리초"],

        // ─── Italian (it) ─────────────────────────────────────────
        ("it", "years", "long") => &["", "anno", "", "", "", "anni"],
        ("it", "years", _) => &["", "a", "", "", "", "a"],
        ("it", "months", "long") => &["", "mese", "", "", "", "mesi"],
        ("it", "months", _) => &["", "m", "", "", "", "m"],
        ("it", "weeks", "long") => &["", "settimana", "", "", "", "settimane"],
        ("it", "weeks", _) => &["", "sett.", "", "", "", "sett."],
        ("it", "days", "long") => &["", "giorno", "", "", "", "giorni"],
        ("it", "days", _) => &["", "g", "", "", "", "g"],
        ("it", "hours", "long") => &["", "ora", "", "", "", "ore"],
        ("it", "hours", _) => &["", "h", "", "", "", "h"],
        ("it", "minutes", "long") => &["", "minuto", "", "", "", "minuti"],
        ("it", "minutes", _) => &["", "min", "", "", "", "min"],
        ("it", "seconds", "long") => &["", "secondo", "", "", "", "secondi"],
        ("it", "seconds", _) => &["", "s", "", "", "", "s"],
        ("it", "milliseconds", "long") => &["", "millisecondo", "", "", "", "millisecondi"],
        ("it", "milliseconds", _) => &["", "ms", "", "", "", "ms"],

        // ─── Turkish (tr) ─────────────────────────────────────────
        ("tr", "years", "long") => &["", "yıl", "", "", "", "yıl"],
        ("tr", "months", "long") => &["", "ay", "", "", "", "ay"],
        ("tr", "weeks", "long") => &["", "hafta", "", "", "", "hafta"],
        ("tr", "days", "long") => &["", "gün", "", "", "", "gün"],
        ("tr", "hours", "long") => &["", "saat", "", "", "", "saat"],
        ("tr", "minutes", "long") => &["", "dakika", "", "", "", "dakika"],
        ("tr", "seconds", "long") => &["", "saniye", "", "", "", "saniye"],
        ("tr", "milliseconds", "long") => &["", "milisaniye", "", "", "", "milisaniye"],

        // ─── Vietnamese (vi) — no plural inflection ───────────────
        ("vi", "years", "long") => &["", "năm", "", "", "", "năm"],
        ("vi", "months", "long") => &["", "tháng", "", "", "", "tháng"],
        ("vi", "weeks", "long") => &["", "tuần", "", "", "", "tuần"],
        ("vi", "days", "long") => &["", "ngày", "", "", "", "ngày"],
        ("vi", "hours", "long") => &["", "giờ", "", "", "", "giờ"],
        ("vi", "minutes", "long") => &["", "phút", "", "", "", "phút"],
        ("vi", "seconds", "long") => &["", "giây", "", "", "", "giây"],
        ("vi", "milliseconds", "long") => &["", "mili giây", "", "", "", "mili giây"],

        // ─── Polish (pl) — 3-way plural ──────────────────────────
        // Forms ordered: [_, one, _, few, many, other]
        // one: 1   few: 2-4 (excluding 12-14)   many: 0, 5+, 12-14
        ("pl", "years", "long") => &["", "rok", "", "lata", "lat", "lat"],
        ("pl", "years", _) => &["", "r.", "", "r.", "r.", "r."],
        ("pl", "months", "long") => &["", "miesiąc", "", "miesiące", "miesięcy", "miesięcy"],
        ("pl", "months", _) => &["", "mies.", "", "mies.", "mies.", "mies."],
        ("pl", "weeks", "long") => &["", "tydzień", "", "tygodnie", "tygodni", "tygodni"],
        ("pl", "weeks", _) => &["", "tyg.", "", "tyg.", "tyg.", "tyg."],
        ("pl", "days", "long") => &["", "dzień", "", "dni", "dni", "dni"],
        ("pl", "days", _) => &["", "dz.", "", "dz.", "dz.", "dz."],
        ("pl", "hours", "long") => &["", "godzina", "", "godziny", "godzin", "godzin"],
        ("pl", "hours", _) => &["", "godz.", "", "godz.", "godz.", "godz."],
        ("pl", "minutes", "long") => &["", "minuta", "", "minuty", "minut", "minut"],
        ("pl", "minutes", _) => &["", "min", "", "min", "min", "min"],
        ("pl", "seconds", "long") => &["", "sekunda", "", "sekundy", "sekund", "sekund"],
        ("pl", "seconds", _) => &["", "s", "", "s", "s", "s"],

        // ─── Indonesian (id) — no plural inflection ──────────────
        ("id", "years", "long") => &["", "tahun", "", "", "", "tahun"],
        ("id", "months", "long") => &["", "bulan", "", "", "", "bulan"],
        ("id", "weeks", "long") => &["", "minggu", "", "", "", "minggu"],
        ("id", "days", "long") => &["", "hari", "", "", "", "hari"],
        ("id", "hours", "long") => &["", "jam", "", "", "", "jam"],
        ("id", "minutes", "long") => &["", "menit", "", "", "", "menit"],
        ("id", "seconds", "long") => &["", "detik", "", "", "", "detik"],
        ("id", "milliseconds", "long") => &["", "milidetik", "", "", "", "milidetik"],

        // ─── Dutch (nl) ───────────────────────────────────────────
        ("nl", "years", "long") => &["", "jaar", "", "", "", "jaar"],
        ("nl", "months", "long") => &["", "maand", "", "", "", "maanden"],
        ("nl", "weeks", "long") => &["", "week", "", "", "", "weken"],
        ("nl", "days", "long") => &["", "dag", "", "", "", "dagen"],
        ("nl", "hours", "long") => &["", "uur", "", "", "", "uur"],
        ("nl", "minutes", "long") => &["", "minuut", "", "", "", "minuten"],
        ("nl", "seconds", "long") => &["", "seconde", "", "", "", "seconden"],
        ("nl", "milliseconds", "long") => &["", "milliseconde", "", "", "", "milliseconden"],

        // ─── Thai (th) — no plural inflection ────────────────────
        ("th", "years", "long") => &["", "ปี", "", "", "", "ปี"],
        ("th", "months", "long") => &["", "เดือน", "", "", "", "เดือน"],
        ("th", "weeks", "long") => &["", "สัปดาห์", "", "", "", "สัปดาห์"],
        ("th", "days", "long") => &["", "วัน", "", "", "", "วัน"],
        ("th", "hours", "long") => &["", "ชั่วโมง", "", "", "", "ชั่วโมง"],
        ("th", "minutes", "long") => &["", "นาที", "", "", "", "นาที"],
        ("th", "seconds", "long") => &["", "วินาที", "", "", "", "วินาที"],
        ("th", "milliseconds", "long") => &["", "มิลลิวินาที", "", "", "", "มิลลิวินาที"],

        // ─── Swedish (sv) ─────────────────────────────────────────
        ("sv", "years", "long") => &["", "år", "", "", "", "år"],
        ("sv", "months", "long") => &["", "månad", "", "", "", "månader"],
        ("sv", "weeks", "long") => &["", "vecka", "", "", "", "veckor"],
        ("sv", "days", "long") => &["", "dag", "", "", "", "dagar"],
        ("sv", "hours", "long") => &["", "timme", "", "", "", "timmar"],
        ("sv", "minutes", "long") => &["", "minut", "", "", "", "minuter"],
        ("sv", "seconds", "long") => &["", "sekund", "", "", "", "sekunder"],
        ("sv", "milliseconds", "long") => &["", "millisekund", "", "", "", "millisekunder"],

        _ => return None,
    })
}

// ── Intl static methods ──────────────────────────────────────────────

fn register_static(vm: &mut VM) {
    vm.register_host_fn(
        "ecma:intl",
        "getCanonicalLocales",
        Box::new(|_ctx, args| {
            let tags: Vec<String> = match args.first() {
                Some(Value::String(s)) => vec![s.to_string()],
                Some(Value::Object(o)) => {
                    let lock = o.lock().unwrap();
                    if let ObjectKind::Array(ref elems) = lock.kind {
                        elems
                            .iter()
                            .map(|v| match v {
                                Value::String(s) => s.to_string(),
                                o => format!("{}", o),
                            })
                            .collect()
                    } else {
                        Vec::new()
                    }
                }
                _ => Vec::new(),
            };
            let canon: Vec<Value> = tags
                .iter()
                .map(|t| {
                    let langid = parse_langid(t);
                    s_val(&langid.to_string())
                })
                .collect();
            make_array(canon)
        }),
    );

    vm.register_host_fn(
        "ecma:intl",
        "supportedValuesOf",
        Box::new(|_ctx, args| {
            let key = match args.first() {
                Some(Value::String(s)) => s.to_string(),
                _ => return make_array(vec![]),
            };
            // Time zones come from tzdb, not a hand-written list. The previous
            // 15-entry list advertised zones that `Intl.DateTimeFormat` would
            // then reject, and tzdb ships 5–10 updates a year, so any literal
            // table here is wrong by construction.
            if key == "timeZone" {
                let mut names: Vec<&str> =
                    chrono_tz::TZ_VARIANTS.iter().map(|tz| tz.name()).collect();
                names.sort_unstable();
                return make_array(
                    names
                        .into_iter()
                        .map(|n| Value::String(Arc::from(n)))
                        .collect(),
                );
            }
            let values: Vec<&'static str> = match key.as_str() {
                "calendar" => vec![
                    "gregory", "buddhist", "chinese", "coptic", "ethiopic", "ethioaa", "hebrew",
                    "indian", "islamic", "iso8601", "japanese", "persian", "roc",
                ],
                "collation" => vec![
                    "compat", "dict", "ducet", "emoji", "eor", "phonebk", "phonetic", "pinyin",
                    "reformed", "searchjl", "stroke", "trad", "unihan", "zhuyin",
                ],
                "currency" => vec![
                    "USD", "EUR", "GBP", "JPY", "CNY", "AUD", "CAD", "CHF", "HKD", "INR", "KRW",
                    "MXN", "NZD", "SGD",
                ],
                "numberingSystem" => vec![
                    "arab", "arabext", "bali", "beng", "deva", "fullwide", "gujr", "guru",
                    "hanidec", "hant", "khmr", "knda", "laoo", "latn", "limb", "mlym", "mong",
                    "mymr", "orya", "tamldec", "telu", "thai", "tibt",
                ],
                "unit" => vec![
                    "acre",
                    "bit",
                    "byte",
                    "celsius",
                    "centimeter",
                    "day",
                    "degree",
                    "fahrenheit",
                    "fluid-ounce",
                    "foot",
                    "gallon",
                    "gigabit",
                    "gigabyte",
                    "gram",
                    "hectare",
                    "hour",
                    "inch",
                    "kilobit",
                    "kilobyte",
                    "kilogram",
                    "kilometer",
                    "liter",
                    "megabit",
                    "megabyte",
                    "meter",
                    "microsecond",
                    "mile",
                    "mile-scandinavian",
                    "milliliter",
                    "millimeter",
                    "millisecond",
                    "minute",
                    "month",
                    "nanosecond",
                    "ounce",
                    "percent",
                    "petabyte",
                    "pound",
                    "second",
                    "stone",
                    "terabit",
                    "terabyte",
                    "week",
                    "yard",
                    "year",
                ],
                _ => vec![],
            };
            make_array(values.into_iter().map(s_val).collect())
        }),
    );
}

#[allow(dead_code)]
fn _force_host_context_use(_: &mut HostContext) {}
