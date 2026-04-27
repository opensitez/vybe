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
use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::{HostContext, VM};

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

fn make_array(elements: Vec<Value>) -> Value {
    Value::Object(Arc::new(Mutex::new(Object::new_array(elements))))
}

fn make_object(props: Vec<(&str, Value)>) -> Value {
    let mut obj = Object::new();
    for (k, v) in props {
        obj.properties.insert(k.into(), v);
    }
    Value::Object(Arc::new(Mutex::new(obj)))
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
    match arg {
        Some(Value::String(tag)) if !tag.is_empty() => tag.to_string(),
        Some(Value::Object(o)) => {
            let lock = o.lock().unwrap();
            if let ObjectKind::Array(ref elems) = lock.kind {
                if let Some(Value::String(tag)) = elems.first() {
                    return tag.to_string();
                }
            }
            "en-US".into()
        }
        _ => "en-US".into(),
    }
}

fn resolve_options(arg: Option<&Value>) -> Arc<Mutex<Object>> {
    if let Some(Value::Object(o)) = arg {
        o.clone()
    } else {
        Arc::new(Mutex::new(Object::new()))
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

fn register_collator(vm: &mut VM) {
    vm.register_host_fn("ecma:intl/collator", "new", Box::new(|_ctx, args| {
        let locale = resolve_locale(args.first());
        let options = resolve_options(args.get(1));
        let opts_lock = options.lock().unwrap();
        let usage = opts_lock.properties.get("usage")
            .and_then(|v| if let Value::String(s) = v { Some(s.to_string()) } else { None })
            .unwrap_or_else(|| "sort".into());
        let sensitivity = opts_lock.properties.get("sensitivity")
            .and_then(|v| if let Value::String(s) = v { Some(s.to_string()) } else { None })
            .unwrap_or_else(|| "variant".into());
        drop(opts_lock);
        make_object(vec![
            ("__type", s_val("Collator")),
            ("locale", s_val(&locale)),
            ("usage", s_val(&usage)),
            ("sensitivity", s_val(&sensitivity)),
        ])
    }));

    vm.register_host_fn("ecma:intl/collator", "compare", Box::new(|_ctx, args| {
        use icu::collator::{Collator, options::CollatorOptions};
        let collator = match args.first() {
            Some(Value::Object(o)) => o.clone(),
            _ => return Value::I32(0),
        };
        let a = match args.get(1) {
            Some(Value::String(s)) => s.to_string(),
            Some(o) => format!("{}", o),
            None => String::new(),
        };
        let b = match args.get(2) {
            Some(Value::String(s)) => s.to_string(),
            Some(o) => format!("{}", o),
            None => String::new(),
        };
        let locale = obj_string_prop(&collator, "locale").unwrap_or_else(|| "en-US".into());
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
    }));

    vm.register_host_fn("ecma:intl/collator", "resolvedOptions", Box::new(|_ctx, args| {
        if let Some(Value::Object(c)) = args.first() {
            let locale = obj_string_prop(c, "locale").unwrap_or_else(|| "en-US".into());
            let usage = obj_string_prop(c, "usage").unwrap_or_else(|| "sort".into());
            let sensitivity = obj_string_prop(c, "sensitivity").unwrap_or_else(|| "variant".into());
            return make_object(vec![
                ("locale", s_val(&locale)),
                ("usage", s_val(&usage)),
                ("sensitivity", s_val(&sensitivity)),
                ("ignorePunctuation", Value::Bool(false)),
                ("collation", s_val("default")),
                ("numeric", Value::Bool(false)),
                ("caseFirst", s_val("false")),
            ]);
        }
        make_object(vec![])
    }));

    vm.register_host_fn("ecma:intl/collator", "supportedLocalesOf", Box::new(|_ctx, args| {
        match args.first() {
            Some(v @ Value::Object(_)) => v.clone(),
            Some(Value::String(s)) => make_array(vec![Value::String(s.clone())]),
            _ => make_array(vec![]),
        }
    }));
}

// ── Intl.NumberFormat (ECMA-402 §15) ─────────────────────────────────

fn register_number_format(vm: &mut VM) {
    vm.register_host_fn("ecma:intl/numberformat", "new", Box::new(|_ctx, args| {
        let locale = resolve_locale(args.first());
        let options = resolve_options(args.get(1));
        let ol = options.lock().unwrap();
        let style = ol.properties.get("style")
            .and_then(|v| if let Value::String(s) = v { Some(s.to_string()) } else { None })
            .unwrap_or_else(|| "decimal".into());
        let currency = ol.properties.get("currency")
            .and_then(|v| if let Value::String(s) = v { Some(s.to_string()) } else { None })
            .unwrap_or_default();
        let min_frac = ol.properties.get("minimumFractionDigits")
            .map(|v| v.as_i32()).unwrap_or(if style == "currency" { 2 } else { 0 });
        let max_frac = ol.properties.get("maximumFractionDigits")
            .map(|v| v.as_i32()).unwrap_or(if style == "currency" { 2 } else { 3 });
        drop(ol);
        make_object(vec![
            ("__type", s_val("NumberFormat")),
            ("locale", s_val(&locale)),
            ("style", s_val(&style)),
            ("currency", s_val(&currency)),
            ("minimumFractionDigits", Value::I32(min_frac)),
            ("maximumFractionDigits", Value::I32(max_frac)),
        ])
    }));

    vm.register_host_fn("ecma:intl/numberformat", "format", Box::new(|_ctx, args| {
        let nf = match args.first() {
            Some(Value::Object(o)) => o.clone(),
            _ => return s_val(""),
        };
        let value = match args.get(1) {
            Some(v) => v.as_f64(),
            None => f64::NAN,
        };
        s_val(&format_number_real(&nf, value))
    }));

    vm.register_host_fn("ecma:intl/numberformat", "formatToParts", Box::new(|_ctx, args| {
        let nf = match args.first() {
            Some(Value::Object(o)) => o.clone(),
            _ => return make_array(vec![]),
        };
        let value = match args.get(1) {
            Some(v) => v.as_f64(),
            None => f64::NAN,
        };
        // MVP: single { type: "literal", value: format(value) } part.
        // Full breakdown into integer/decimal/group/currency parts is
        // a future enhancement (icu::decimal exposes part info).
        make_array(vec![make_object(vec![
            ("type", s_val("literal")),
            ("value", s_val(&format_number_real(&nf, value))),
        ])])
    }));

    vm.register_host_fn("ecma:intl/numberformat", "resolvedOptions", Box::new(|_ctx, args| {
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
    }));

    vm.register_host_fn("ecma:intl/numberformat", "supportedLocalesOf", Box::new(|_ctx, args| {
        match args.first() {
            Some(v @ Value::Object(_)) => v.clone(),
            Some(Value::String(s)) => make_array(vec![Value::String(s.clone())]),
            _ => make_array(vec![]),
        }
    }));
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
        lock.properties.get("maximumFractionDigits").map(|v| v.as_i32()).unwrap_or(3)
    };

    if value.is_nan() { return "NaN".into(); }
    if value.is_infinite() { return if value < 0.0 { "-∞".into() } else { "∞".into() }; }

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
    let scaled_value = if style == "percent" { value * 100.0 } else { value };
    let mult = 10f64.powi(max_frac);
    let int_form = (scaled_value * mult).round() as i64;
    let mut fd: Decimal = int_form.into();
    if max_frac > 0 {
        fd.multiply_pow10(-(max_frac as i16));
    }
    let min_frac = {
        let lock = nf.lock().unwrap();
        lock.properties.get("minimumFractionDigits").map(|v| v.as_i32()).unwrap_or(0)
    };
    fd.trim_end();
    // After trim, pad back up to min_frac if needed (e.g. currency
    // wants always 2 decimals).
    if min_frac > 0 {
        fd.pad_end(-(min_frac as i16));
    }
    let formatted = formatter.format(&fd).to_string();

    match style.as_str() {
        "currency" => {
            let symbol = currency_symbol(&currency);
            format!("{}{}", symbol, formatted)
        }
        "percent" => format!("{}%", formatted),
        _ => formatted,
    }
}

fn currency_symbol(code: &str) -> &'static str {
    match code {
        "USD" => "$", "EUR" => "€", "GBP" => "£", "JPY" => "¥",
        "CHF" => "CHF ", "CAD" | "AUD" | "NZD" => "$",
        _ => "¤",
    }
}

// ── Intl.DateTimeFormat (ECMA-402 §13) ───────────────────────────────

fn register_date_time_format(vm: &mut VM) {
    vm.register_host_fn("ecma:intl/datetimeformat", "new", Box::new(|_ctx, args| {
        let locale = resolve_locale(args.first());
        let _options = resolve_options(args.get(1));
        make_object(vec![
            ("__type", s_val("DateTimeFormat")),
            ("locale", s_val(&locale)),
        ])
    }));

    vm.register_host_fn("ecma:intl/datetimeformat", "format", Box::new(|_ctx, args| {
        let dtf = match args.first() {
            Some(Value::Object(o)) => o.clone(),
            _ => return s_val(""),
        };
        let ms = match args.get(1) {
            Some(v) => v.as_f64(),
            None => return s_val(""),
        };
        s_val(&format_date_real(&dtf, ms))
    }));

    vm.register_host_fn("ecma:intl/datetimeformat", "formatToParts", Box::new(|_ctx, args| {
        let dtf = match args.first() {
            Some(Value::Object(o)) => o.clone(),
            _ => return make_array(vec![]),
        };
        let ms = match args.get(1) {
            Some(v) => v.as_f64(),
            None => return make_array(vec![]),
        };
        make_array(vec![make_object(vec![
            ("type", s_val("literal")),
            ("value", s_val(&format_date_real(&dtf, ms))),
        ])])
    }));

    vm.register_host_fn("ecma:intl/datetimeformat", "resolvedOptions", Box::new(|_ctx, args| {
        if let Some(Value::Object(dtf)) = args.first() {
            let locale = obj_string_prop(dtf, "locale").unwrap_or_else(|| "en-US".into());
            return make_object(vec![
                ("locale", s_val(&locale)),
                ("calendar", s_val("gregory")),
                ("numberingSystem", s_val("latn")),
                ("timeZone", s_val("UTC")),
            ]);
        }
        make_object(vec![])
    }));

    vm.register_host_fn("ecma:intl/datetimeformat", "supportedLocalesOf", Box::new(|_ctx, args| {
        match args.first() {
            Some(v @ Value::Object(_)) => v.clone(),
            Some(Value::String(s)) => make_array(vec![Value::String(s.clone())]),
            _ => make_array(vec![]),
        }
    }));
}

/// Format ms-since-epoch into a locale-aware date string using
/// `icu::datetime::DateTimeFormatter`. Uses the YMD::medium() field
/// set per ECMA-402 §13.1.2 default behaviour (year, month, day —
/// no time without explicit options).
fn format_date_real(dtf: &Arc<Mutex<Object>>, ms: f64) -> String {
    use icu::datetime::{DateTimeFormatter, fieldsets};
    use icu::datetime::input::Date;

    let locale = obj_string_prop(dtf, "locale").unwrap_or_else(|| "en-US".into());
    let icu_loc = parse_icu_locale(&locale);

    let secs = (ms / 1000.0) as i64;
    let (y, mo, d) = epoch_to_ymd(secs);

    let date = match Date::try_new_iso(y, mo as u8, d as u8) {
        Ok(d) => d,
        Err(_) => return String::new(),
    };
    let formatter = match DateTimeFormatter::try_new(
        (&icu_loc).into(),
        fieldsets::YMD::medium(),
    ) {
        Ok(f) => f,
        Err(_) => return format!("{}/{}/{}", mo, d, y),
    };
    formatter.format(&date).to_string()
}

fn epoch_to_ymd(secs: i64) -> (i32, i32, i32) {
    let days = secs.div_euclid(86400);
    let mut year = 1970i32;
    let mut remaining = days;
    loop {
        let leap = is_leap(year);
        let year_days = if leap { 366 } else { 365 };
        if remaining < year_days as i64 { break; }
        remaining -= year_days as i64;
        year += 1;
    }
    let leap = is_leap(year);
    let months = [31, if leap { 29 } else { 28 }, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    let mut month = 1;
    for &m in &months {
        if remaining < m as i64 { break; }
        remaining -= m as i64;
        month += 1;
    }
    let day = remaining as i32 + 1;
    (year, month, day)
}

fn is_leap(y: i32) -> bool {
    (y % 4 == 0 && y % 100 != 0) || y % 400 == 0
}

// ── Intl.ListFormat (ECMA-402 §14) ───────────────────────────────────

fn register_list_format(vm: &mut VM) {
    vm.register_host_fn("ecma:intl/listformat", "new", Box::new(|_ctx, args| {
        let locale = resolve_locale(args.first());
        let options = resolve_options(args.get(1));
        let ol = options.lock().unwrap();
        let lf_type = ol.properties.get("type")
            .and_then(|v| if let Value::String(s) = v { Some(s.to_string()) } else { None })
            .unwrap_or_else(|| "conjunction".into());
        let style = ol.properties.get("style")
            .and_then(|v| if let Value::String(s) = v { Some(s.to_string()) } else { None })
            .unwrap_or_else(|| "long".into());
        drop(ol);
        make_object(vec![
            ("__type", s_val("ListFormat")),
            ("locale", s_val(&locale)),
            ("type", s_val(&lf_type)),
            ("style", s_val(&style)),
        ])
    }));

    vm.register_host_fn("ecma:intl/listformat", "format", Box::new(|_ctx, args| {
        use icu::list::{ListFormatter, options::{ListFormatterOptions, ListLength}};
        let lf = match args.first() {
            Some(Value::Object(o)) => o.clone(),
            _ => return s_val(""),
        };
        let items: Vec<String> = match args.get(1) {
            Some(Value::Object(o)) => {
                let lock = o.lock().unwrap();
                if let ObjectKind::Array(ref elems) = lock.kind {
                    elems.iter().map(|v| match v { Value::String(s) => s.to_string(), o => format!("{}", o) }).collect()
                } else { Vec::new() }
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
            "unit"        => ListFormatter::try_new_unit(prefs, opts),
            _             => ListFormatter::try_new_and(prefs, opts),
        };
        let formatter = match formatter {
            Ok(f) => f,
            Err(_) => return s_val(&fallback_format_list(&items, &lf_type)),
        };
        s_val(&formatter.format_to_string(items.iter()))
    }));

    vm.register_host_fn("ecma:intl/listformat", "formatToParts", Box::new(|_ctx, args| {
        let lf = match args.first() {
            Some(Value::Object(o)) => o.clone(),
            _ => return make_array(vec![]),
        };
        let items: Vec<String> = match args.get(1) {
            Some(Value::Object(o)) => {
                let lock = o.lock().unwrap();
                if let ObjectKind::Array(ref elems) = lock.kind {
                    elems.iter().map(|v| match v { Value::String(s) => s.to_string(), o => format!("{}", o) }).collect()
                } else { Vec::new() }
            }
            _ => Vec::new(),
        };
        let lf_type = obj_string_prop(&lf, "type").unwrap_or_else(|| "conjunction".into());
        // MVP: single literal part; full breakdown is a future enhancement.
        make_array(vec![make_object(vec![
            ("type", s_val("literal")),
            ("value", s_val(&fallback_format_list(&items, &lf_type))),
        ])])
    }));

    vm.register_host_fn("ecma:intl/listformat", "resolvedOptions", Box::new(|_ctx, args| {
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
    }));

    vm.register_host_fn("ecma:intl/listformat", "supportedLocalesOf", Box::new(|_ctx, args| {
        match args.first() {
            Some(v @ Value::Object(_)) => v.clone(),
            Some(Value::String(s)) => make_array(vec![Value::String(s.clone())]),
            _ => make_array(vec![]),
        }
    }));
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
    vm.register_host_fn("ecma:intl/pluralrules", "new", Box::new(|_ctx, args| {
        let locale = resolve_locale(args.first());
        let options = resolve_options(args.get(1));
        let ol = options.lock().unwrap();
        let pr_type = ol.properties.get("type")
            .and_then(|v| if let Value::String(s) = v { Some(s.to_string()) } else { None })
            .unwrap_or_else(|| "cardinal".into());
        drop(ol);
        make_object(vec![
            ("__type", s_val("PluralRules")),
            ("locale", s_val(&locale)),
            ("type", s_val(&pr_type)),
        ])
    }));

    vm.register_host_fn("ecma:intl/pluralrules", "select", Box::new(|_ctx, args| {
        let pr = match args.first() {
            Some(Value::Object(o)) => o.clone(),
            _ => return s_val("other"),
        };
        let n = match args.get(1) {
            Some(v) => v.as_f64(),
            None => 0.0,
        };
        let locale = obj_string_prop(&pr, "locale").unwrap_or_else(|| "en-US".into());
        let pr_type = obj_string_prop(&pr, "type").unwrap_or_else(|| "cardinal".into());
        s_val(plural_select_real(&locale, &pr_type, n))
    }));

    vm.register_host_fn("ecma:intl/pluralrules", "selectRange", Box::new(|_ctx, args| {
        let pr = match args.first() {
            Some(Value::Object(o)) => o.clone(),
            _ => return s_val("other"),
        };
        // ECMA-402 selectRange returns the plural category for the
        // formatted range. Without locale-specific range rules, return
        // the end value's category (good enough approximation).
        let end = args.get(2).map(|v| v.as_f64()).unwrap_or(0.0);
        let locale = obj_string_prop(&pr, "locale").unwrap_or_else(|| "en-US".into());
        let pr_type = obj_string_prop(&pr, "type").unwrap_or_else(|| "cardinal".into());
        s_val(plural_select_real(&locale, &pr_type, end))
    }));

    vm.register_host_fn("ecma:intl/pluralrules", "resolvedOptions", Box::new(|_ctx, args| {
        if let Some(Value::Object(pr)) = args.first() {
            let locale = obj_string_prop(pr, "locale").unwrap_or_else(|| "en-US".into());
            let pr_type = obj_string_prop(pr, "type").unwrap_or_else(|| "cardinal".into());
            return make_object(vec![
                ("locale", s_val(&locale)),
                ("type", s_val(&pr_type)),
                ("minimumIntegerDigits", Value::I32(1)),
                ("minimumFractionDigits", Value::I32(0)),
                ("maximumFractionDigits", Value::I32(3)),
                ("pluralCategories", make_array(vec![
                    s_val("zero"), s_val("one"), s_val("two"),
                    s_val("few"), s_val("many"), s_val("other"),
                ])),
            ]);
        }
        make_object(vec![])
    }));

    vm.register_host_fn("ecma:intl/pluralrules", "supportedLocalesOf", Box::new(|_ctx, args| {
        match args.first() {
            Some(v @ Value::Object(_)) => v.clone(),
            Some(Value::String(s)) => make_array(vec![Value::String(s.clone())]),
            _ => make_array(vec![]),
        }
    }));
}

fn plural_select_real(locale: &str, pr_type: &str, n: f64) -> &'static str {
    use intl_pluralrules::{PluralCategory, PluralRuleType, PluralRules};
    // intl_pluralrules indexes rules by BASE LANGUAGE only ("en", "fr",
    // ...). Strip region/script so "en-US" still finds the "en" table.
    let parsed = parse_langid(locale);
    let langid = unic_langid::LanguageIdentifier::from_parts(
        parsed.language, None, None, &[],
    );
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
    vm.register_host_fn("ecma:intl/relativetimeformat", "new", Box::new(|_ctx, args| {
        let locale = resolve_locale(args.first());
        let options = resolve_options(args.get(1));
        let ol = options.lock().unwrap();
        let numeric = ol.properties.get("numeric")
            .and_then(|v| if let Value::String(s) = v { Some(s.to_string()) } else { None })
            .unwrap_or_else(|| "always".into());
        let style = ol.properties.get("style")
            .and_then(|v| if let Value::String(s) = v { Some(s.to_string()) } else { None })
            .unwrap_or_else(|| "long".into());
        drop(ol);
        make_object(vec![
            ("__type", s_val("RelativeTimeFormat")),
            ("locale", s_val(&locale)),
            ("numeric", s_val(&numeric)),
            ("style", s_val(&style)),
        ])
    }));

    vm.register_host_fn("ecma:intl/relativetimeformat", "format", Box::new(|_ctx, args| {
        let rtf = match args.first() {
            Some(Value::Object(o)) => o.clone(),
            _ => return s_val(""),
        };
        let value = match args.get(1) { Some(v) => v.as_f64(), None => 0.0 };
        let unit = match args.get(2) { Some(Value::String(s)) => s.to_string(), _ => "second".into() };
        let locale = obj_string_prop(&rtf, "locale").unwrap_or_else(|| "en-US".into());
        let style = obj_string_prop(&rtf, "style").unwrap_or_else(|| "long".into());
        s_val(&format_relative_time_real(&locale, &style, value, &unit))
    }));

    vm.register_host_fn("ecma:intl/relativetimeformat", "formatToParts", Box::new(|_ctx, args| {
        let rtf = match args.first() {
            Some(Value::Object(o)) => o.clone(),
            _ => return make_array(vec![]),
        };
        let value = match args.get(1) { Some(v) => v.as_f64(), None => 0.0 };
        let unit = match args.get(2) { Some(Value::String(s)) => s.to_string(), _ => "second".into() };
        let locale = obj_string_prop(&rtf, "locale").unwrap_or_else(|| "en-US".into());
        let style = obj_string_prop(&rtf, "style").unwrap_or_else(|| "long".into());
        make_array(vec![make_object(vec![
            ("type", s_val("literal")),
            ("value", s_val(&format_relative_time_real(&locale, &style, value, &unit))),
        ])])
    }));

    vm.register_host_fn("ecma:intl/relativetimeformat", "resolvedOptions", Box::new(|_ctx, args| {
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
    }));

    vm.register_host_fn("ecma:intl/relativetimeformat", "supportedLocalesOf", Box::new(|_ctx, args| {
        match args.first() {
            Some(v @ Value::Object(_)) => v.clone(),
            Some(Value::String(s)) => make_array(vec![Value::String(s.clone())]),
            _ => make_array(vec![]),
        }
    }));
}

/// Format relative time using `icu_relativetime`. Constructor variants
/// per (style × unit) — pick the right one and format. Falls back to
/// a plain English form on locale/unit data miss.
fn format_relative_time_real(locale: &str, style: &str, value: f64, unit: &str) -> String {
    use icu_relativetime::{RelativeTimeFormatter, RelativeTimeFormatterOptions};
    use icu_relativetime::options::Numeric;
    use fixed_decimal::FixedDecimal;

    // icu_relativetime 0.1 uses the older icu_locid types — separate
    // from icu 2.x's icu_locale_core. Bridge via the older crate.
    let icu_loc: icu_locid::Locale = match locale.parse() {
        Ok(l) => l,
        Err(_) => return fallback_relative_time(value, unit),
    };

    let opts = RelativeTimeFormatterOptions { numeric: Numeric::Always };
    let unit_norm = unit.trim_end_matches('s'); // "days" → "day"

    let formatter_result = match (style, unit_norm) {
        ("short", "second")  => RelativeTimeFormatter::try_new_short_second(&icu_loc.into(), opts),
        ("short", "minute")  => RelativeTimeFormatter::try_new_short_minute(&icu_loc.into(), opts),
        ("short", "hour")    => RelativeTimeFormatter::try_new_short_hour(&icu_loc.into(), opts),
        ("short", "day")     => RelativeTimeFormatter::try_new_short_day(&icu_loc.into(), opts),
        ("short", "week")    => RelativeTimeFormatter::try_new_short_week(&icu_loc.into(), opts),
        ("short", "month")   => RelativeTimeFormatter::try_new_short_month(&icu_loc.into(), opts),
        ("short", "quarter") => RelativeTimeFormatter::try_new_short_quarter(&icu_loc.into(), opts),
        ("short", "year")    => RelativeTimeFormatter::try_new_short_year(&icu_loc.into(), opts),
        ("narrow", "second")  => RelativeTimeFormatter::try_new_narrow_second(&icu_loc.into(), opts),
        ("narrow", "minute")  => RelativeTimeFormatter::try_new_narrow_minute(&icu_loc.into(), opts),
        ("narrow", "hour")    => RelativeTimeFormatter::try_new_narrow_hour(&icu_loc.into(), opts),
        ("narrow", "day")     => RelativeTimeFormatter::try_new_narrow_day(&icu_loc.into(), opts),
        ("narrow", "week")    => RelativeTimeFormatter::try_new_narrow_week(&icu_loc.into(), opts),
        ("narrow", "month")   => RelativeTimeFormatter::try_new_narrow_month(&icu_loc.into(), opts),
        ("narrow", "quarter") => RelativeTimeFormatter::try_new_narrow_quarter(&icu_loc.into(), opts),
        ("narrow", "year")    => RelativeTimeFormatter::try_new_narrow_year(&icu_loc.into(), opts),
        // Default style "long" or unknown.
        (_, "second")  => RelativeTimeFormatter::try_new_long_second(&icu_loc.into(), opts),
        (_, "minute")  => RelativeTimeFormatter::try_new_long_minute(&icu_loc.into(), opts),
        (_, "hour")    => RelativeTimeFormatter::try_new_long_hour(&icu_loc.into(), opts),
        (_, "day")     => RelativeTimeFormatter::try_new_long_day(&icu_loc.into(), opts),
        (_, "week")    => RelativeTimeFormatter::try_new_long_week(&icu_loc.into(), opts),
        (_, "month")   => RelativeTimeFormatter::try_new_long_month(&icu_loc.into(), opts),
        (_, "quarter") => RelativeTimeFormatter::try_new_long_quarter(&icu_loc.into(), opts),
        (_, "year")    => RelativeTimeFormatter::try_new_long_year(&icu_loc.into(), opts),
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

fn fallback_relative_time(value: f64, unit: &str) -> String {
    let abs = value.abs() as i64;
    let stem = unit.trim_end_matches('s');
    let unit_str = if abs == 1 { stem.to_string() } else { format!("{}s", stem) };
    if value < 0.0 {
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
    vm.register_host_fn("ecma:intl/segmenter", "new", Box::new(|_ctx, args| {
        let locale = resolve_locale(args.first());
        let options = resolve_options(args.get(1));
        let ol = options.lock().unwrap();
        let granularity = ol.properties.get("granularity")
            .and_then(|v| if let Value::String(s) = v { Some(s.to_string()) } else { None })
            .unwrap_or_else(|| "grapheme".into());
        drop(ol);
        make_object(vec![
            ("__type", s_val("Segmenter")),
            ("locale", s_val(&locale)),
            ("granularity", s_val(&granularity)),
        ])
    }));

    vm.register_host_fn("ecma:intl/segmenter", "segment", Box::new(|_ctx, args| {
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
        let granularity = obj_string_prop(&seg, "granularity").unwrap_or_else(|| "grapheme".into());

        let segments: Vec<(String, usize)> = match granularity.as_str() {
            "word" => input.split_word_bound_indices()
                .map(|(i, s)| (s.to_string(), i))
                .collect(),
            "sentence" => input.split_sentence_bound_indices()
                .map(|(i, s)| (s.to_string(), i))
                .collect(),
            _ => input.grapheme_indices(true)
                .map(|(i, s)| (s.to_string(), i))
                .collect(),
        };

        let elems: Vec<Value> = segments.into_iter().map(|(seg_str, idx)| {
            make_object(vec![
                ("segment", s_val(&seg_str)),
                ("index", Value::I32(idx as i32)),
                ("input", s_val(&input)),
            ])
        }).collect();
        make_array(elems)
    }));

    vm.register_host_fn("ecma:intl/segmenter", "resolvedOptions", Box::new(|_ctx, args| {
        if let Some(Value::Object(seg)) = args.first() {
            let locale = obj_string_prop(seg, "locale").unwrap_or_else(|| "en-US".into());
            let granularity = obj_string_prop(seg, "granularity").unwrap_or_else(|| "grapheme".into());
            return make_object(vec![
                ("locale", s_val(&locale)),
                ("granularity", s_val(&granularity)),
            ]);
        }
        make_object(vec![])
    }));

    vm.register_host_fn("ecma:intl/segmenter", "supportedLocalesOf", Box::new(|_ctx, args| {
        match args.first() {
            Some(v @ Value::Object(_)) => v.clone(),
            Some(Value::String(s)) => make_array(vec![Value::String(s.clone())]),
            _ => make_array(vec![]),
        }
    }));
}

// ── Intl.Locale (ECMA-402 §14) ───────────────────────────────────────
//
// Backed by `unic-langid` for parsing — produces canonical BCP-47
// representation with proper subtag normalization.

fn register_locale(vm: &mut VM) {
    vm.register_host_fn("ecma:intl/locale", "new", Box::new(|_ctx, args| {
        let tag = match args.first() {
            Some(Value::String(s)) => s.to_string(),
            Some(o) => format!("{}", o),
            None => "en".into(),
        };
        let langid = parse_langid(&tag);
        let language = langid.language.as_str().to_string();
        let region = langid.region.map(|r| r.as_str().to_string()).unwrap_or_default();
        let script = langid.script.map(|s| s.as_str().to_string()).unwrap_or_default();
        let base_name = langid.to_string();
        make_object(vec![
            ("__type", s_val("Locale")),
            ("baseName", s_val(&base_name)),
            ("language", s_val(&language)),
            ("region", s_val(&region)),
            ("script", s_val(&script)),
            ("calendar", s_val("")),
            ("numberingSystem", s_val("")),
            ("collation", s_val("")),
            ("caseFirst", s_val("")),
            ("hourCycle", s_val("")),
            ("numeric", Value::Bool(false)),
        ])
    }));

    vm.register_host_fn("ecma:intl/locale", "toString", Box::new(|_ctx, args| {
        if let Some(Value::Object(loc)) = args.first() {
            return s_val(&obj_string_prop(loc, "baseName").unwrap_or_default());
        }
        s_val("")
    }));

    vm.register_host_fn("ecma:intl/locale", "maximize", Box::new(|_ctx, args| {
        // Likely-subtag expansion needs CLDR data; unic-langid 0.9
        // alone doesn't provide it. Return as-is — full expansion is
        // a future enhancement (could use icu::locale::Locale::maximize).
        args.first().cloned().unwrap_or(Value::Null)
    }));

    vm.register_host_fn("ecma:intl/locale", "minimize", Box::new(|_ctx, args| {
        args.first().cloned().unwrap_or(Value::Null)
    }));
}

// ── Intl.DisplayNames (ECMA-402 §12) ─────────────────────────────────
//
// Backed by `icu_displaynames` — full CLDR translation tables.

fn register_display_names(vm: &mut VM) {
    vm.register_host_fn("ecma:intl/displaynames", "new", Box::new(|_ctx, args| {
        let locale = resolve_locale(args.first());
        let options = resolve_options(args.get(1));
        let ol = options.lock().unwrap();
        let dn_type = ol.properties.get("type")
            .and_then(|v| if let Value::String(s) = v { Some(s.to_string()) } else { None })
            .unwrap_or_else(|| "language".into());
        let style = ol.properties.get("style")
            .and_then(|v| if let Value::String(s) = v { Some(s.to_string()) } else { None })
            .unwrap_or_else(|| "long".into());
        drop(ol);
        make_object(vec![
            ("__type", s_val("DisplayNames")),
            ("locale", s_val(&locale)),
            ("type", s_val(&dn_type)),
            ("style", s_val(&style)),
        ])
    }));

    vm.register_host_fn("ecma:intl/displaynames", "of", Box::new(|_ctx, args| {
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
    }));

    vm.register_host_fn("ecma:intl/displaynames", "resolvedOptions", Box::new(|_ctx, args| {
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
    }));

    vm.register_host_fn("ecma:intl/displaynames", "supportedLocalesOf", Box::new(|_ctx, args| {
        match args.first() {
            Some(v @ Value::Object(_)) => v.clone(),
            Some(Value::String(s)) => make_array(vec![Value::String(s.clone())]),
            _ => make_array(vec![]),
        }
    }));
}

/// Return the display name for `code` in the given locale's perspective.
/// `dn_type` is "language" / "region" / "script" / "currency".
///
/// `icu_displaynames` 0.11 has the API but its locale parsing API uses
/// the older `icu_locid` (transitive dep). We bridge by parsing the
/// tag through that crate.
fn display_name_of(locale: &str, dn_type: &str, code: &str) -> String {
    use icu_displaynames::{DisplayNamesOptions, LanguageDisplayNames, RegionDisplayNames, ScriptDisplayNames};

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
    vm.register_host_fn("ecma:intl/durationformat", "new", Box::new(|_ctx, args| {
        let locale = resolve_locale(args.first());
        let options = resolve_options(args.get(1));
        let ol = options.lock().unwrap();
        let style = ol.properties.get("style")
            .and_then(|v| if let Value::String(s) = v { Some(s.to_string()) } else { None })
            .unwrap_or_else(|| "short".into());
        drop(ol);
        make_object(vec![
            ("__type", s_val("DurationFormat")),
            ("locale", s_val(&locale)),
            ("style", s_val(&style)),
        ])
    }));

    vm.register_host_fn("ecma:intl/durationformat", "format", Box::new(|_ctx, args| {
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
    }));

    vm.register_host_fn("ecma:intl/durationformat", "formatToParts", Box::new(|_ctx, args| {
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
    }));

    vm.register_host_fn("ecma:intl/durationformat", "resolvedOptions", Box::new(|_ctx, args| {
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
    }));

    vm.register_host_fn("ecma:intl/durationformat", "supportedLocalesOf", Box::new(|_ctx, args| {
        match args.first() {
            Some(v @ Value::Object(_)) => v.clone(),
            Some(Value::String(s)) => make_array(vec![Value::String(s.clone())]),
            _ => make_array(vec![]),
        }
    }));
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
    let langid = unic_langid::LanguageIdentifier::from_parts(
        parsed.language, None, None, &[],
    );
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
        let h = lock.properties.get("hours").map(|v| v.as_i32()).unwrap_or(0);
        let m = lock.properties.get("minutes").map(|v| v.as_i32()).unwrap_or(0);
        let s = lock.properties.get("seconds").map(|v| v.as_i32()).unwrap_or(0);
        let mut clock = format!("{}:{:02}:{:02}", h, m.abs(), s.abs());
        let mut extras = Vec::new();
        for key in &["years", "months", "weeks", "days", "milliseconds", "microseconds", "nanoseconds"] {
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
    for key in &["years", "months", "weeks", "days", "hours", "minutes", "seconds", "milliseconds", "microseconds", "nanoseconds"] {
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
fn duration_unit_text(lang: &str, unit: &str, style: &str, category: intl_pluralrules::PluralCategory) -> &'static str {
    use intl_pluralrules::PluralCategory::*;
    // Table layout: each (lang, unit, style) entry is an array of forms
    // ordered by category index. We use a closure for lookup so the
    // match is contiguous and the table stays readable.
    let cat_idx = match category {
        ZERO => 0, ONE => 1, TWO => 2, FEW => 3, MANY => 4, OTHER => 5,
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
        ("en", "years",        "long")   => &["", "year", "", "", "", "years"],
        ("en", "years",        "short")  => &["", "yr", "", "", "", "yrs"],
        ("en", "years",        "narrow") => &["", "y", "", "", "", "y"],
        ("en", "months",       "long")   => &["", "month", "", "", "", "months"],
        ("en", "months",       "short")  => &["", "mo", "", "", "", "mo"],
        ("en", "months",       "narrow") => &["", "mo", "", "", "", "mo"],
        ("en", "weeks",        "long")   => &["", "week", "", "", "", "weeks"],
        ("en", "weeks",        "short")  => &["", "wk", "", "", "", "wks"],
        ("en", "weeks",        "narrow") => &["", "w", "", "", "", "w"],
        ("en", "days",         "long")   => &["", "day", "", "", "", "days"],
        ("en", "days",         "short")  => &["", "day", "", "", "", "days"],
        ("en", "days",         "narrow") => &["", "d", "", "", "", "d"],
        ("en", "hours",        "long")   => &["", "hour", "", "", "", "hours"],
        ("en", "hours",        "short")  => &["", "hr", "", "", "", "hr"],
        ("en", "hours",        "narrow") => &["", "h", "", "", "", "h"],
        ("en", "minutes",      "long")   => &["", "minute", "", "", "", "minutes"],
        ("en", "minutes",      "short")  => &["", "min", "", "", "", "min"],
        ("en", "minutes",      "narrow") => &["", "m", "", "", "", "m"],
        ("en", "seconds",      "long")   => &["", "second", "", "", "", "seconds"],
        ("en", "seconds",      "short")  => &["", "sec", "", "", "", "sec"],
        ("en", "seconds",      "narrow") => &["", "s", "", "", "", "s"],
        ("en", "milliseconds", "long")   => &["", "millisecond", "", "", "", "milliseconds"],
        ("en", "milliseconds", "short")  => &["", "ms", "", "", "", "ms"],
        ("en", "milliseconds", "narrow") => &["", "ms", "", "", "", "ms"],
        ("en", "microseconds", "long")   => &["", "microsecond", "", "", "", "microseconds"],
        ("en", "microseconds", "short")  => &["", "μs", "", "", "", "μs"],
        ("en", "microseconds", "narrow") => &["", "μs", "", "", "", "μs"],
        ("en", "nanoseconds",  "long")   => &["", "nanosecond", "", "", "", "nanoseconds"],
        ("en", "nanoseconds",  "short")  => &["", "ns", "", "", "", "ns"],
        ("en", "nanoseconds",  "narrow") => &["", "ns", "", "", "", "ns"],

        // ─── Mandarin Chinese (zh) — no plural inflection ─────────
        ("zh", "years", _)        => &["", "年", "", "", "", "年"],
        ("zh", "months", _)       => &["", "个月", "", "", "", "个月"],
        ("zh", "weeks", _)        => &["", "周", "", "", "", "周"],
        ("zh", "days", _)         => &["", "天", "", "", "", "天"],
        ("zh", "hours", _)        => &["", "小时", "", "", "", "小时"],
        ("zh", "minutes", _)      => &["", "分钟", "", "", "", "分钟"],
        ("zh", "seconds", _)      => &["", "秒", "", "", "", "秒"],
        ("zh", "milliseconds", _) => &["", "毫秒", "", "", "", "毫秒"],
        ("zh", "microseconds", _) => &["", "微秒", "", "", "", "微秒"],
        ("zh", "nanoseconds", _)  => &["", "纳秒", "", "", "", "纳秒"],

        // ─── Spanish (es) ─────────────────────────────────────────
        ("es", "years",        "long")   => &["", "año", "", "", "", "años"],
        ("es", "years",        "short")  => &["", "a", "", "", "", "a"],
        ("es", "years",        "narrow") => &["", "a", "", "", "", "a"],
        ("es", "months",       "long")   => &["", "mes", "", "", "", "meses"],
        ("es", "months",       _)        => &["", "m", "", "", "", "m"],
        ("es", "weeks",        "long")   => &["", "semana", "", "", "", "semanas"],
        ("es", "weeks",        _)        => &["", "sem.", "", "", "", "sem."],
        ("es", "days",         "long")   => &["", "día", "", "", "", "días"],
        ("es", "days",         _)        => &["", "d", "", "", "", "d"],
        ("es", "hours",        "long")   => &["", "hora", "", "", "", "horas"],
        ("es", "hours",        _)        => &["", "h", "", "", "", "h"],
        ("es", "minutes",      "long")   => &["", "minuto", "", "", "", "minutos"],
        ("es", "minutes",      _)        => &["", "min", "", "", "", "min"],
        ("es", "seconds",      "long")   => &["", "segundo", "", "", "", "segundos"],
        ("es", "seconds",      _)        => &["", "s", "", "", "", "s"],
        ("es", "milliseconds", "long")   => &["", "milisegundo", "", "", "", "milisegundos"],
        ("es", "milliseconds", _)        => &["", "ms", "", "", "", "ms"],
        ("es", "microseconds", "long")   => &["", "microsegundo", "", "", "", "microsegundos"],
        ("es", "microseconds", _)        => &["", "μs", "", "", "", "μs"],
        ("es", "nanoseconds",  "long")   => &["", "nanosegundo", "", "", "", "nanosegundos"],
        ("es", "nanoseconds",  _)        => &["", "ns", "", "", "", "ns"],

        // ─── Hindi (hi) ───────────────────────────────────────────
        ("hi", "years",        "long")   => &["", "वर्ष", "", "", "", "वर्ष"],
        ("hi", "months",       "long")   => &["", "महीना", "", "", "", "महीने"],
        ("hi", "weeks",        "long")   => &["", "सप्ताह", "", "", "", "सप्ताह"],
        ("hi", "days",         "long")   => &["", "दिन", "", "", "", "दिन"],
        ("hi", "hours",        "long")   => &["", "घंटा", "", "", "", "घंटे"],
        ("hi", "minutes",      "long")   => &["", "मिनट", "", "", "", "मिनट"],
        ("hi", "seconds",      "long")   => &["", "सेकंड", "", "", "", "सेकंड"],
        ("hi", "milliseconds", "long")   => &["", "मिलीसेकंड", "", "", "", "मिलीसेकंड"],
        ("hi", "microseconds", "long")   => &["", "माइक्रोसेकंड", "", "", "", "माइक्रोसेकंड"],
        ("hi", "nanoseconds",  "long")   => &["", "नैनोसेकंड", "", "", "", "नैनोसेकंड"],

        // ─── Arabic (ar) — 6 plural categories ────────────────────
        // CLDR forms: zero, one, two, few (3-10), many (11+), other.
        // Modern Standard Arabic uses different word forms per category.
        ("ar", "years", "long") => &[
            "سنة",       // zero (لا توجد سنوات)
            "سنة",       // one (سنة واحدة)
            "سنتان",     // two (سنتان)
            "سنوات",     // few (3-10 sanawat)
            "سنة",       // many (11+ "sana")
            "سنة",       // other (decimals)
        ],
        ("ar", "years", _) => &["", "سنة", "", "", "", "سنة"],
        ("ar", "months", "long") => &[
            "شهر", "شهر", "شهران", "أشهر", "شهرًا", "شهر",
        ],
        ("ar", "months", _) => &["", "شهر", "", "", "", "شهر"],
        ("ar", "weeks", "long") => &[
            "أسبوع", "أسبوع", "أسبوعان", "أسابيع", "أسبوعًا", "أسبوع",
        ],
        ("ar", "weeks", _) => &["", "أسبوع", "", "", "", "أسبوع"],
        ("ar", "days", "long") => &[
            "يوم", "يوم", "يومان", "أيام", "يومًا", "يوم",
        ],
        ("ar", "days", _) => &["", "يوم", "", "", "", "يوم"],
        ("ar", "hours", "long") => &[
            "ساعة", "ساعة", "ساعتان", "ساعات", "ساعة", "ساعة",
        ],
        ("ar", "hours", _) => &["", "س", "", "", "", "س"],
        ("ar", "minutes", "long") => &[
            "دقيقة", "دقيقة", "دقيقتان", "دقائق", "دقيقة", "دقيقة",
        ],
        ("ar", "minutes", _) => &["", "د", "", "", "", "د"],
        ("ar", "seconds", "long") => &[
            "ثانية", "ثانية", "ثانيتان", "ثوانٍ", "ثانية", "ثانية",
        ],
        ("ar", "seconds", _) => &["", "ث", "", "", "", "ث"],
        ("ar", "milliseconds", "long") => &["", "ميلي ثانية", "", "", "", "ميلي ثانية"],
        ("ar", "milliseconds", _) => &["", "ms", "", "", "", "ms"],
        ("ar", "microseconds", "long") => &["", "ميكروثانية", "", "", "", "ميكروثانية"],
        ("ar", "microseconds", _) => &["", "μs", "", "", "", "μs"],
        ("ar", "nanoseconds", "long") => &["", "نانوثانية", "", "", "", "نانوثانية"],
        ("ar", "nanoseconds", _) => &["", "ns", "", "", "", "ns"],

        // ─── Portuguese (pt) ──────────────────────────────────────
        ("pt", "years",        "long")   => &["", "ano", "", "", "", "anos"],
        ("pt", "years",        _)        => &["", "a", "", "", "", "a"],
        ("pt", "months",       "long")   => &["", "mês", "", "", "", "meses"],
        ("pt", "months",       _)        => &["", "m", "", "", "", "m"],
        ("pt", "weeks",        "long")   => &["", "semana", "", "", "", "semanas"],
        ("pt", "weeks",        _)        => &["", "sem.", "", "", "", "sem."],
        ("pt", "days",         "long")   => &["", "dia", "", "", "", "dias"],
        ("pt", "days",         _)        => &["", "d", "", "", "", "d"],
        ("pt", "hours",        "long")   => &["", "hora", "", "", "", "horas"],
        ("pt", "hours",        _)        => &["", "h", "", "", "", "h"],
        ("pt", "minutes",      "long")   => &["", "minuto", "", "", "", "minutos"],
        ("pt", "minutes",      _)        => &["", "min", "", "", "", "min"],
        ("pt", "seconds",      "long")   => &["", "segundo", "", "", "", "segundos"],
        ("pt", "seconds",      _)        => &["", "s", "", "", "", "s"],
        ("pt", "milliseconds", "long")   => &["", "milissegundo", "", "", "", "milissegundos"],
        ("pt", "milliseconds", _)        => &["", "ms", "", "", "", "ms"],

        // ─── Bengali (bn) ─────────────────────────────────────────
        ("bn", "years",        "long")   => &["", "বছর", "", "", "", "বছর"],
        ("bn", "months",       "long")   => &["", "মাস", "", "", "", "মাস"],
        ("bn", "weeks",        "long")   => &["", "সপ্তাহ", "", "", "", "সপ্তাহ"],
        ("bn", "days",         "long")   => &["", "দিন", "", "", "", "দিন"],
        ("bn", "hours",        "long")   => &["", "ঘণ্টা", "", "", "", "ঘণ্টা"],
        ("bn", "minutes",      "long")   => &["", "মিনিট", "", "", "", "মিনিট"],
        ("bn", "seconds",      "long")   => &["", "সেকেন্ড", "", "", "", "সেকেন্ড"],

        // ─── Russian (ru) — 3-way plural ──────────────────────────
        // Forms ordered: [_, one, _, few, many, other]
        // one: 1, 21, 31...   few: 2-4, 22-24...   many: 0, 5-20...
        ("ru", "years", "long")   => &["", "год", "", "года", "лет", "года"],
        ("ru", "years", _)        => &["", "г.", "", "г.", "г.", "г."],
        ("ru", "months", "long")  => &["", "месяц", "", "месяца", "месяцев", "месяца"],
        ("ru", "months", _)       => &["", "мес.", "", "мес.", "мес.", "мес."],
        ("ru", "weeks", "long")   => &["", "неделя", "", "недели", "недель", "недели"],
        ("ru", "weeks", _)        => &["", "нед.", "", "нед.", "нед.", "нед."],
        ("ru", "days", "long")    => &["", "день", "", "дня", "дней", "дня"],
        ("ru", "days", _)         => &["", "дн.", "", "дн.", "дн.", "дн."],
        ("ru", "hours", "long")   => &["", "час", "", "часа", "часов", "часа"],
        ("ru", "hours", _)        => &["", "ч.", "", "ч.", "ч.", "ч."],
        ("ru", "minutes", "long") => &["", "минута", "", "минуты", "минут", "минуты"],
        ("ru", "minutes", _)      => &["", "мин.", "", "мин.", "мин.", "мин."],
        ("ru", "seconds", "long") => &["", "секунда", "", "секунды", "секунд", "секунды"],
        ("ru", "seconds", _)      => &["", "сек.", "", "сек.", "сек.", "сек."],
        ("ru", "milliseconds", "long") => &["", "миллисекунда", "", "миллисекунды", "миллисекунд", "миллисекунды"],
        ("ru", "milliseconds", _) => &["", "мс", "", "мс", "мс", "мс"],

        // ─── Japanese (ja) — no plural inflection ─────────────────
        ("ja", "years", _)        => &["", "年", "", "", "", "年"],
        ("ja", "months", _)       => &["", "か月", "", "", "", "か月"],
        ("ja", "weeks", _)        => &["", "週間", "", "", "", "週間"],
        ("ja", "days", _)         => &["", "日", "", "", "", "日"],
        ("ja", "hours", _)        => &["", "時間", "", "", "", "時間"],
        ("ja", "minutes", _)      => &["", "分", "", "", "", "分"],
        ("ja", "seconds", _)      => &["", "秒", "", "", "", "秒"],
        ("ja", "milliseconds", _) => &["", "ミリ秒", "", "", "", "ミリ秒"],
        ("ja", "microseconds", _) => &["", "マイクロ秒", "", "", "", "マイクロ秒"],
        ("ja", "nanoseconds", _)  => &["", "ナノ秒", "", "", "", "ナノ秒"],

        // ─── German (de) ──────────────────────────────────────────
        ("de", "years",        "long")   => &["", "Jahr", "", "", "", "Jahre"],
        ("de", "years",        _)        => &["", "J", "", "", "", "J"],
        ("de", "months",       "long")   => &["", "Monat", "", "", "", "Monate"],
        ("de", "months",       _)        => &["", "Mon.", "", "", "", "Mon."],
        ("de", "weeks",        "long")   => &["", "Woche", "", "", "", "Wochen"],
        ("de", "weeks",        _)        => &["", "Wo.", "", "", "", "Wo."],
        ("de", "days",         "long")   => &["", "Tag", "", "", "", "Tage"],
        ("de", "days",         _)        => &["", "Tg.", "", "", "", "Tg."],
        ("de", "hours",        "long")   => &["", "Stunde", "", "", "", "Stunden"],
        ("de", "hours",        _)        => &["", "Std.", "", "", "", "Std."],
        ("de", "minutes",      "long")   => &["", "Minute", "", "", "", "Minuten"],
        ("de", "minutes",      _)        => &["", "Min.", "", "", "", "Min."],
        ("de", "seconds",      "long")   => &["", "Sekunde", "", "", "", "Sekunden"],
        ("de", "seconds",      _)        => &["", "Sek.", "", "", "", "Sek."],
        ("de", "milliseconds", "long")   => &["", "Millisekunde", "", "", "", "Millisekunden"],
        ("de", "milliseconds", _)        => &["", "ms", "", "", "", "ms"],

        // ─── French (fr) — `one` covers 0 and 1 ───────────────────
        ("fr", "years",        "long")   => &["", "an", "", "", "", "ans"],
        ("fr", "years",        _)        => &["", "an", "", "", "", "ans"],
        ("fr", "months",       "long")   => &["", "mois", "", "", "", "mois"],
        ("fr", "months",       _)        => &["", "m.", "", "", "", "m."],
        ("fr", "weeks",        "long")   => &["", "semaine", "", "", "", "semaines"],
        ("fr", "weeks",        _)        => &["", "sem.", "", "", "", "sem."],
        ("fr", "days",         "long")   => &["", "jour", "", "", "", "jours"],
        ("fr", "days",         _)        => &["", "j", "", "", "", "j"],
        ("fr", "hours",        "long")   => &["", "heure", "", "", "", "heures"],
        ("fr", "hours",        _)        => &["", "h", "", "", "", "h"],
        ("fr", "minutes",      "long")   => &["", "minute", "", "", "", "minutes"],
        ("fr", "minutes",      _)        => &["", "min", "", "", "", "min"],
        ("fr", "seconds",      "long")   => &["", "seconde", "", "", "", "secondes"],
        ("fr", "seconds",      _)        => &["", "s", "", "", "", "s"],
        ("fr", "milliseconds", "long")   => &["", "milliseconde", "", "", "", "millisecondes"],
        ("fr", "milliseconds", _)        => &["", "ms", "", "", "", "ms"],

        // ─── Korean (ko) — no plural inflection ───────────────────
        ("ko", "years", _)        => &["", "년", "", "", "", "년"],
        ("ko", "months", _)       => &["", "개월", "", "", "", "개월"],
        ("ko", "weeks", _)        => &["", "주", "", "", "", "주"],
        ("ko", "days", _)         => &["", "일", "", "", "", "일"],
        ("ko", "hours", _)        => &["", "시간", "", "", "", "시간"],
        ("ko", "minutes", _)      => &["", "분", "", "", "", "분"],
        ("ko", "seconds", _)      => &["", "초", "", "", "", "초"],
        ("ko", "milliseconds", _) => &["", "밀리초", "", "", "", "밀리초"],

        // ─── Italian (it) ─────────────────────────────────────────
        ("it", "years",        "long")   => &["", "anno", "", "", "", "anni"],
        ("it", "years",        _)        => &["", "a", "", "", "", "a"],
        ("it", "months",       "long")   => &["", "mese", "", "", "", "mesi"],
        ("it", "months",       _)        => &["", "m", "", "", "", "m"],
        ("it", "weeks",        "long")   => &["", "settimana", "", "", "", "settimane"],
        ("it", "weeks",        _)        => &["", "sett.", "", "", "", "sett."],
        ("it", "days",         "long")   => &["", "giorno", "", "", "", "giorni"],
        ("it", "days",         _)        => &["", "g", "", "", "", "g"],
        ("it", "hours",        "long")   => &["", "ora", "", "", "", "ore"],
        ("it", "hours",        _)        => &["", "h", "", "", "", "h"],
        ("it", "minutes",      "long")   => &["", "minuto", "", "", "", "minuti"],
        ("it", "minutes",      _)        => &["", "min", "", "", "", "min"],
        ("it", "seconds",      "long")   => &["", "secondo", "", "", "", "secondi"],
        ("it", "seconds",      _)        => &["", "s", "", "", "", "s"],
        ("it", "milliseconds", "long")   => &["", "millisecondo", "", "", "", "millisecondi"],
        ("it", "milliseconds", _)        => &["", "ms", "", "", "", "ms"],

        // ─── Turkish (tr) ─────────────────────────────────────────
        ("tr", "years",        "long")   => &["", "yıl", "", "", "", "yıl"],
        ("tr", "months",       "long")   => &["", "ay", "", "", "", "ay"],
        ("tr", "weeks",        "long")   => &["", "hafta", "", "", "", "hafta"],
        ("tr", "days",         "long")   => &["", "gün", "", "", "", "gün"],
        ("tr", "hours",        "long")   => &["", "saat", "", "", "", "saat"],
        ("tr", "minutes",      "long")   => &["", "dakika", "", "", "", "dakika"],
        ("tr", "seconds",      "long")   => &["", "saniye", "", "", "", "saniye"],
        ("tr", "milliseconds", "long")   => &["", "milisaniye", "", "", "", "milisaniye"],

        // ─── Vietnamese (vi) — no plural inflection ───────────────
        ("vi", "years",        "long")   => &["", "năm", "", "", "", "năm"],
        ("vi", "months",       "long")   => &["", "tháng", "", "", "", "tháng"],
        ("vi", "weeks",        "long")   => &["", "tuần", "", "", "", "tuần"],
        ("vi", "days",         "long")   => &["", "ngày", "", "", "", "ngày"],
        ("vi", "hours",        "long")   => &["", "giờ", "", "", "", "giờ"],
        ("vi", "minutes",      "long")   => &["", "phút", "", "", "", "phút"],
        ("vi", "seconds",      "long")   => &["", "giây", "", "", "", "giây"],
        ("vi", "milliseconds", "long")   => &["", "mili giây", "", "", "", "mili giây"],

        // ─── Polish (pl) — 3-way plural ──────────────────────────
        // Forms ordered: [_, one, _, few, many, other]
        // one: 1   few: 2-4 (excluding 12-14)   many: 0, 5+, 12-14
        ("pl", "years", "long")   => &["", "rok", "", "lata", "lat", "lat"],
        ("pl", "years", _)        => &["", "r.", "", "r.", "r.", "r."],
        ("pl", "months", "long")  => &["", "miesiąc", "", "miesiące", "miesięcy", "miesięcy"],
        ("pl", "months", _)       => &["", "mies.", "", "mies.", "mies.", "mies."],
        ("pl", "weeks", "long")   => &["", "tydzień", "", "tygodnie", "tygodni", "tygodni"],
        ("pl", "weeks", _)        => &["", "tyg.", "", "tyg.", "tyg.", "tyg."],
        ("pl", "days", "long")    => &["", "dzień", "", "dni", "dni", "dni"],
        ("pl", "days", _)         => &["", "dz.", "", "dz.", "dz.", "dz."],
        ("pl", "hours", "long")   => &["", "godzina", "", "godziny", "godzin", "godzin"],
        ("pl", "hours", _)        => &["", "godz.", "", "godz.", "godz.", "godz."],
        ("pl", "minutes", "long") => &["", "minuta", "", "minuty", "minut", "minut"],
        ("pl", "minutes", _)      => &["", "min", "", "min", "min", "min"],
        ("pl", "seconds", "long") => &["", "sekunda", "", "sekundy", "sekund", "sekund"],
        ("pl", "seconds", _)      => &["", "s", "", "s", "s", "s"],

        // ─── Indonesian (id) — no plural inflection ──────────────
        ("id", "years",        "long")   => &["", "tahun", "", "", "", "tahun"],
        ("id", "months",       "long")   => &["", "bulan", "", "", "", "bulan"],
        ("id", "weeks",        "long")   => &["", "minggu", "", "", "", "minggu"],
        ("id", "days",         "long")   => &["", "hari", "", "", "", "hari"],
        ("id", "hours",        "long")   => &["", "jam", "", "", "", "jam"],
        ("id", "minutes",      "long")   => &["", "menit", "", "", "", "menit"],
        ("id", "seconds",      "long")   => &["", "detik", "", "", "", "detik"],
        ("id", "milliseconds", "long")   => &["", "milidetik", "", "", "", "milidetik"],

        // ─── Dutch (nl) ───────────────────────────────────────────
        ("nl", "years",        "long")   => &["", "jaar", "", "", "", "jaar"],
        ("nl", "months",       "long")   => &["", "maand", "", "", "", "maanden"],
        ("nl", "weeks",        "long")   => &["", "week", "", "", "", "weken"],
        ("nl", "days",         "long")   => &["", "dag", "", "", "", "dagen"],
        ("nl", "hours",        "long")   => &["", "uur", "", "", "", "uur"],
        ("nl", "minutes",      "long")   => &["", "minuut", "", "", "", "minuten"],
        ("nl", "seconds",      "long")   => &["", "seconde", "", "", "", "seconden"],
        ("nl", "milliseconds", "long")   => &["", "milliseconde", "", "", "", "milliseconden"],

        // ─── Thai (th) — no plural inflection ────────────────────
        ("th", "years",        "long")   => &["", "ปี", "", "", "", "ปี"],
        ("th", "months",       "long")   => &["", "เดือน", "", "", "", "เดือน"],
        ("th", "weeks",        "long")   => &["", "สัปดาห์", "", "", "", "สัปดาห์"],
        ("th", "days",         "long")   => &["", "วัน", "", "", "", "วัน"],
        ("th", "hours",        "long")   => &["", "ชั่วโมง", "", "", "", "ชั่วโมง"],
        ("th", "minutes",      "long")   => &["", "นาที", "", "", "", "นาที"],
        ("th", "seconds",      "long")   => &["", "วินาที", "", "", "", "วินาที"],
        ("th", "milliseconds", "long")   => &["", "มิลลิวินาที", "", "", "", "มิลลิวินาที"],

        // ─── Swedish (sv) ─────────────────────────────────────────
        ("sv", "years",        "long")   => &["", "år", "", "", "", "år"],
        ("sv", "months",       "long")   => &["", "månad", "", "", "", "månader"],
        ("sv", "weeks",        "long")   => &["", "vecka", "", "", "", "veckor"],
        ("sv", "days",         "long")   => &["", "dag", "", "", "", "dagar"],
        ("sv", "hours",        "long")   => &["", "timme", "", "", "", "timmar"],
        ("sv", "minutes",      "long")   => &["", "minut", "", "", "", "minuter"],
        ("sv", "seconds",      "long")   => &["", "sekund", "", "", "", "sekunder"],
        ("sv", "milliseconds", "long")   => &["", "millisekund", "", "", "", "millisekunder"],

        _ => return None,
    })
}

// ── Intl static methods ──────────────────────────────────────────────

fn register_static(vm: &mut VM) {
    vm.register_host_fn("ecma:intl", "getCanonicalLocales", Box::new(|_ctx, args| {
        let tags: Vec<String> = match args.first() {
            Some(Value::String(s)) => vec![s.to_string()],
            Some(Value::Object(o)) => {
                let lock = o.lock().unwrap();
                if let ObjectKind::Array(ref elems) = lock.kind {
                    elems.iter().map(|v| match v { Value::String(s) => s.to_string(), o => format!("{}", o) }).collect()
                } else { Vec::new() }
            }
            _ => Vec::new(),
        };
        let canon: Vec<Value> = tags.iter().map(|t| {
            let langid = parse_langid(t);
            s_val(&langid.to_string())
        }).collect();
        make_array(canon)
    }));

    vm.register_host_fn("ecma:intl", "supportedValuesOf", Box::new(|_ctx, args| {
        let key = match args.first() {
            Some(Value::String(s)) => s.to_string(),
            _ => return make_array(vec![]),
        };
        let values: Vec<&'static str> = match key.as_str() {
            "calendar" => vec!["gregory", "buddhist", "chinese", "coptic", "ethiopic", "ethioaa", "hebrew", "indian", "islamic", "iso8601", "japanese", "persian", "roc"],
            "collation" => vec!["compat", "dict", "ducet", "emoji", "eor", "phonebk", "phonetic", "pinyin", "reformed", "searchjl", "stroke", "trad", "unihan", "zhuyin"],
            "currency" => vec!["USD", "EUR", "GBP", "JPY", "CNY", "AUD", "CAD", "CHF", "HKD", "INR", "KRW", "MXN", "NZD", "SGD"],
            "numberingSystem" => vec!["arab", "arabext", "bali", "beng", "deva", "fullwide", "gujr", "guru", "hanidec", "hant", "khmr", "knda", "laoo", "latn", "limb", "mlym", "mong", "mymr", "orya", "tamldec", "telu", "thai", "tibt"],
            "timeZone" => vec!["UTC", "America/New_York", "America/Los_Angeles", "America/Chicago", "America/Denver", "Europe/London", "Europe/Paris", "Europe/Berlin", "Europe/Moscow", "Asia/Tokyo", "Asia/Shanghai", "Asia/Hong_Kong", "Asia/Singapore", "Asia/Dubai", "Australia/Sydney"],
            "unit" => vec!["acre", "bit", "byte", "celsius", "centimeter", "day", "degree", "fahrenheit", "fluid-ounce", "foot", "gallon", "gigabit", "gigabyte", "gram", "hectare", "hour", "inch", "kilobit", "kilobyte", "kilogram", "kilometer", "liter", "megabit", "megabyte", "meter", "microsecond", "mile", "mile-scandinavian", "milliliter", "millimeter", "millisecond", "minute", "month", "nanosecond", "ounce", "percent", "petabyte", "pound", "second", "stone", "terabit", "terabyte", "week", "yard", "year"],
            _ => vec![],
        };
        make_array(values.into_iter().map(s_val).collect())
    }));
}

#[allow(dead_code)]
fn _force_host_context_use(_: &mut HostContext) {}
