/// Date-based equivalents of Temporal API concepts — construction, arithmetic,
/// comparison, formatting, duration objects, now, and epoch-based instants.
use super::helpers::run_js;

// ── PlainDate equivalents ────────────────────────────────────────────────────

#[test]
fn plain_date_construction() {
    assert_eq!(
        run_js(
            r#"
const d = new Date(2024, 0, 15); // month is 0-indexed
const year = d.getFullYear();
const month = d.getMonth() + 1;
const day = d.getDate();
console.log(year);
console.log(month);
console.log(day);
"#
        ),
        vec!["2024", "1", "15"]
    );
}

#[test]
fn plain_date_from_string() {
    assert_eq!(
        run_js(
            r#"
const parts = "2024-03-21".split("-").map(Number);
const d = new Date(parts[0], parts[1] - 1, parts[2]);
console.log(d.getFullYear());
console.log(d.getMonth() + 1);
console.log(d.getDate());
"#
        ),
        vec!["2024", "3", "21"]
    );
}

#[test]
fn plain_date_tostring() {
    assert_eq!(
        run_js(
            r#"
function pad(n) { return String(n).padStart(2, "0"); }
const d = new Date(2024, 2, 5); // 2024-03-05
const s = d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate());
console.log(s);
"#
        ),
        vec!["2024-03-05"]
    );
}

#[test]
fn plain_date_add_duration() {
    assert_eq!(
        run_js(
            r#"
const d = new Date(2024, 0, 15);
d.setDate(d.getDate() + 10);
console.log(d.getDate());
console.log(d.getMonth() + 1);
"#
        ),
        vec!["25", "1"]
    );
}

#[test]
fn plain_date_add_months() {
    assert_eq!(
        run_js(
            r#"
// Use day 15 to avoid month-overflow edge cases
const d = new Date(2024, 0, 15);
d.setMonth(d.getMonth() + 1);
console.log(d.getMonth() + 1);
"#
        ),
        vec!["2"]
    );
}

#[test]
fn plain_date_subtract() {
    assert_eq!(
        run_js(
            r#"
const d = new Date(2024, 2, 10);
d.setDate(d.getDate() - 5);
console.log(d.getDate());
"#
        ),
        vec!["5"]
    );
}

#[test]
fn plain_date_compare() {
    assert_eq!(
        run_js(
            r#"
function cmp(x, y) { return x < y ? -1 : x > y ? 1 : 0; }
const a = new Date(2024, 0, 1).getTime();
const b = new Date(2024, 5, 1).getTime();
console.log(cmp(a, b));
console.log(cmp(b, a));
console.log(cmp(a, a));
"#
        ),
        vec!["-1", "1", "0"]
    );
}

#[test]
fn plain_date_until_duration() {
    assert_eq!(
        run_js(
            r#"
const start = new Date(2024, 0, 1).getTime();
const end   = new Date(2024, 0, 11).getTime();
const days  = Math.round((end - start) / 86400000);
console.log(days);
"#
        ),
        vec!["10"]
    );
}

#[test]
fn plain_date_since() {
    assert_eq!(
        run_js(
            r#"
const start = new Date(2024, 0, 1).getTime();
const end   = new Date(2024, 0, 11).getTime();
const days  = Math.round((end - start) / 86400000);
console.log(days);
"#
        ),
        vec!["10"]
    );
}

#[test]
fn plain_date_day_of_week() {
    assert_eq!(
        run_js(
            r#"
// Jan 1, 2024 was a Monday; getDay() returns 0=Sun,1=Mon,...
const d = new Date(2024, 0, 1);
console.log(d.getDay()); // 1 = Monday
"#
        ),
        vec!["1"]
    );
}

#[test]
fn plain_date_with_modification() {
    assert_eq!(
        run_js(
            r#"
const orig = new Date(2024, 0, 15);
const d = new Date(orig.getTime());
d.setDate(1);
console.log(d.getDate());
console.log(d.getMonth() + 1);
console.log(d.getFullYear());
"#
        ),
        vec!["1", "1", "2024"]
    );
}

// ── PlainTime equivalents ────────────────────────────────────────────────────

#[test]
fn plain_time_construction() {
    assert_eq!(
        run_js(
            r#"
const d = new Date(2024, 0, 1, 10, 30, 45);
console.log(d.getHours());
console.log(d.getMinutes());
console.log(d.getSeconds());
"#
        ),
        vec!["10", "30", "45"]
    );
}

#[test]
fn plain_time_from_string() {
    assert_eq!(
        run_js(
            r#"
const parts = "14:30:00".split(":").map(Number);
const hour = parts[0], minute = parts[1];
console.log(hour);
console.log(minute);
"#
        ),
        vec!["14", "30"]
    );
}

#[test]
fn plain_time_tostring() {
    assert_eq!(
        run_js(
            r#"
function pad(n) { return String(n).padStart(2, "0"); }
const h = 9, m = 5, s = 3;
console.log(pad(h) + ":" + pad(m) + ":" + pad(s));
"#
        ),
        vec!["09:05:03"]
    );
}

#[test]
fn plain_time_add() {
    assert_eq!(
        run_js(
            r#"
const startH = 10, startM = 30;
const addH = 2, addM = 15;
const totalM = startH * 60 + startM + addH * 60 + addM;
console.log(Math.floor(totalM / 60));
console.log(totalM % 60);
"#
        ),
        vec!["12", "45"]
    );
}

// ── PlainDateTime equivalents ────────────────────────────────────────────────

#[test]
fn plain_datetime_construction() {
    assert_eq!(
        run_js(
            r#"
const dt = new Date(2024, 2, 15, 10, 30, 0);
console.log(dt.getFullYear());
console.log(dt.getMonth() + 1);
console.log(dt.getDate());
console.log(dt.getHours());
console.log(dt.getMinutes());
"#
        ),
        vec!["2024", "3", "15", "10", "30"]
    );
}

#[test]
fn plain_datetime_from_string() {
    assert_eq!(
        run_js(
            r#"
// "2024-03-15T10:30:00" → parse manually
const s = "2024-03-15T10:30:00";
const [datePart, timePart] = s.split("T");
const [y, mo, d] = datePart.split("-").map(Number);
const [h, mi] = timePart.split(":").map(Number);
console.log(y);
console.log(h);
"#
        ),
        vec!["2024", "10"]
    );
}

#[test]
fn plain_datetime_tostring() {
    assert_eq!(
        run_js(
            r#"
function pad(n) { return String(n).padStart(2, "0"); }
const dt = new Date(2024, 2, 5, 9, 5, 3);
const s = `${dt.getFullYear()}-${pad(dt.getMonth()+1)}-${pad(dt.getDate())}T${pad(dt.getHours())}:${pad(dt.getMinutes())}:${pad(dt.getSeconds())}`;
console.log(s);
"#
        ),
        vec!["2024-03-05T09:05:03"]
    );
}

// ── Duration equivalents ─────────────────────────────────────────────────────

#[test]
fn duration_construction() {
    assert_eq!(
        run_js(
            r#"
const dur = { years: 1, months: 2, days: 3, hours: 4, minutes: 5, seconds: 6 };
console.log(dur.years);
console.log(dur.months);
console.log(dur.days);
console.log(dur.hours);
"#
        ),
        vec!["1", "2", "3", "4"]
    );
}

#[test]
fn duration_from_object() {
    assert_eq!(
        run_js(
            r#"
function makeDuration(obj) {
    return { years:0, months:0, days:0, hours:0, minutes:0, seconds:0, ...obj };
}
const dur = makeDuration({ days: 7, hours: 12 });
console.log(dur.days);
console.log(dur.hours);
"#
        ),
        vec!["7", "12"]
    );
}

#[test]
fn duration_negate() {
    assert_eq!(
        run_js(
            r#"
const dur = { days: 5 };
const neg = { days: -dur.days };
console.log(neg.days);
"#
        ),
        vec!["-5"]
    );
}

// ── Now equivalents ──────────────────────────────────────────────────────────

#[test]
fn temporal_now_plaindate_utc() {
    assert_eq!(
        run_js(
            r#"
const year = new Date().getFullYear();
console.log(typeof year === "number");
console.log(year >= 2024);
"#
        ),
        vec!["true", "true"]
    );
}

#[test]
fn temporal_now_instant_has_epoch_seconds() {
    assert_eq!(
        run_js(
            r#"
const epochSeconds = Math.floor(Date.now() / 1000);
console.log(typeof epochSeconds === "number");
console.log(epochSeconds > 1700000000);
"#
        ),
        vec!["true", "true"]
    );
}

// ── Instant equivalents ──────────────────────────────────────────────────────

#[test]
fn instant_from_epoch_milliseconds() {
    assert_eq!(
        run_js(
            r#"
const epochMs = 0;
const epochSeconds = new Date(epochMs).getTime() / 1000;
console.log(epochSeconds);
"#
        ),
        vec!["0"]
    );
}

#[test]
fn instant_compare() {
    assert_eq!(
        run_js(
            r#"
function cmp(x, y) { return x < y ? -1 : x > y ? 1 : 0; }
const a = 1000;
const b = 2000;
console.log(cmp(a, b));
"#
        ),
        vec!["-1"]
    );
}
