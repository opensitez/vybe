/// JavaScript Date object: construction, methods, getters/setters,
/// formatting, arithmetic, comparison, static methods.
use super::helpers::run_js;

// ===================================================================
// DATE CONSTRUCTION
// ===================================================================

#[test]
fn date_now_returns_number() {
    assert_eq!(
        run_js(
            r#"
let ts = Date.now();
console.log(typeof ts);
console.log(ts > 0);
"#
        ),
        &["number", "true"]
    );
}

#[test]
fn date_from_string() {
    assert_eq!(
        run_js(
            r#"
let d = new Date("2024-03-15T00:00:00Z");
console.log(d.getFullYear());
console.log(d.getMonth());
console.log(d.getUTCDate());
"#
        ),
        &["2024", "2", "15"]
    );
}

#[test]
fn date_from_components() {
    assert_eq!(
        run_js(
            r#"
let d = new Date(2024, 0, 15);
console.log(d.getFullYear());
console.log(d.getMonth());
console.log(d.getDate());
"#
        ),
        &["2024", "0", "15"]
    );
}

#[test]
fn date_from_timestamp() {
    assert_eq!(
        run_js(
            r#"
let d = new Date(0);
console.log(d.getUTCFullYear());
console.log(d.getUTCMonth());
console.log(d.getUTCDate());
"#
        ),
        &["1970", "0", "1"]
    );
}

// ===================================================================
// DATE GETTERS
// ===================================================================

#[test]
fn date_getters() {
    assert_eq!(
        run_js(
            r#"
let d = new Date(2024, 5, 15, 10, 30, 45);
console.log(d.getFullYear());
console.log(d.getMonth());
console.log(d.getDate());
console.log(d.getHours());
console.log(d.getMinutes());
console.log(d.getSeconds());
"#
        ),
        &["2024", "5", "15", "10", "30", "45"]
    );
}

#[test]
fn date_get_day() {
    assert_eq!(
        run_js(
            r#"
let d = new Date(2024, 0, 1);
let day = d.getDay();
console.log(day >= 0 && day <= 6);
"#
        ),
        &["true"]
    );
}

#[test]
fn date_get_time() {
    assert_eq!(
        run_js(
            r#"
let d = new Date(0);
console.log(d.getTime());
"#
        ),
        &["0"]
    );
}

// ===================================================================
// DATE SETTERS
// ===================================================================

#[test]
fn date_setters() {
    assert_eq!(
        run_js(
            r#"
let d = new Date(2024, 0, 1);
d.setFullYear(2025);
d.setMonth(11);
d.setDate(25);
console.log(d.getFullYear());
console.log(d.getMonth());
console.log(d.getDate());
"#
        ),
        &["2025", "11", "25"]
    );
}

#[test]
fn date_set_hours_minutes() {
    assert_eq!(
        run_js(
            r#"
let d = new Date(2024, 0, 1, 0, 0, 0);
d.setHours(14);
d.setMinutes(30);
d.setSeconds(59);
console.log(d.getHours());
console.log(d.getMinutes());
console.log(d.getSeconds());
"#
        ),
        &["14", "30", "59"]
    );
}

// ===================================================================
// DATE FORMATTING
// ===================================================================

#[test]
fn date_toisostring() {
    assert_eq!(
        run_js(
            r#"
let d = new Date("2024-03-15T12:00:00Z");
let iso = d.toISOString();
console.log(iso.startsWith("2024-03-15"));
console.log(iso.endsWith("Z"));
"#
        ),
        &["true", "true"]
    );
}

#[test]
fn date_tostring_not_empty() {
    assert_eq!(
        run_js(
            r#"
let d = new Date(2024, 0, 1);
let s = d.toString();
console.log(s.length > 0);
console.log(typeof s);
"#
        ),
        &["true", "string"]
    );
}

#[test]
fn date_todatestring() {
    assert_eq!(
        run_js(
            r#"
let d = new Date(2024, 0, 1);
let s = d.toDateString();
console.log(typeof s);
console.log(s.length > 0);
"#
        ),
        &["string", "true"]
    );
}

#[test]
fn date_json_serialization() {
    assert_eq!(
        run_js(
            r#"
let d = new Date("2024-06-15T00:00:00Z");
let json = JSON.stringify({ date: d });
console.log(json.includes("2024"));
"#
        ),
        &["true"]
    );
}

// ===================================================================
// DATE ARITHMETIC
// ===================================================================

#[test]
fn date_add_days() {
    assert_eq!(
        run_js(
            r#"
let d = new Date(2024, 0, 30);
d.setDate(d.getDate() + 5);
console.log(d.getMonth());
console.log(d.getDate());
"#
        ),
        &["1", "4"]
    );
}

#[test]
fn date_subtract_dates() {
    assert_eq!(
        run_js(
            r#"
let d1 = new Date(2024, 0, 1);
let d2 = new Date(2024, 0, 11);
let diffMs = d2 - d1;
let diffDays = diffMs / (1000 * 60 * 60 * 24);
console.log(diffDays);
"#
        ),
        &["10"]
    );
}

#[test]
fn date_month_overflow() {
    assert_eq!(
        run_js(
            r#"
let d = new Date(2024, 11, 31);
d.setMonth(d.getMonth() + 1);
console.log(d.getFullYear());
console.log(d.getMonth());
"#
        ),
        &["2025", "0"]
    );
}

// ===================================================================
// DATE COMPARISON
// ===================================================================

#[test]
fn date_comparison() {
    assert_eq!(
        run_js(
            r#"
let d1 = new Date(2024, 0, 1);
let d2 = new Date(2024, 6, 1);
console.log(d1 < d2);
console.log(d1 > d2);
console.log(d1.getTime() === new Date(2024, 0, 1).getTime());
"#
        ),
        &["true", "false", "true"]
    );
}

// ===================================================================
// DATE STATIC METHODS
// ===================================================================

#[test]
fn date_parse() {
    assert_eq!(
        run_js(
            r#"
let ts = Date.parse("2024-01-01T00:00:00Z");
console.log(typeof ts);
console.log(ts > 0);
"#
        ),
        &["number", "true"]
    );
}

#[test]
fn date_utc() {
    assert_eq!(
        run_js(
            r#"
let ts = Date.UTC(2024, 0, 1);
let d = new Date(ts);
console.log(d.getUTCFullYear());
console.log(d.getUTCMonth());
console.log(d.getUTCDate());
"#
        ),
        &["2024", "0", "1"]
    );
}
