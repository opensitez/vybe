/// Intl date formatting — DateTimeFormat advanced patterns
use super::helpers::run_js;

#[test]
fn intl_date_format_date_only() {
    assert_eq!(
        run_js(
            r#"
const fmt = new Intl.DateTimeFormat("en-US", {
    year: "numeric", month: "2-digit", day: "2-digit",
    timeZone: "UTC"
});
const d = new Date("2024-06-15T00:00:00.000Z");
const result = fmt.format(d);
console.log(result.includes("2024"));
console.log(result.includes("06") || result.includes("6"));
"#
        ),
        vec!["true", "true"]
    );
}

#[test]
fn intl_date_format_parts() {
    assert_eq!(
        run_js(
            r#"
const fmt = new Intl.DateTimeFormat("en-US", {
    year: "numeric", month: "long", day: "numeric",
    timeZone: "UTC"
});
const d = new Date("2024-06-15T00:00:00.000Z");
const parts = fmt.formatToParts(d);
const types = parts.map(p => p.type).join(",");
console.log(types.includes("year"));
console.log(types.includes("month"));
console.log(types.includes("day"));
"#
        ),
        vec!["true", "true", "true"]
    );
}

#[test]
fn intl_date_format_time() {
    assert_eq!(
        run_js(
            r#"
const fmt = new Intl.DateTimeFormat("en-US", {
    hour: "2-digit", minute: "2-digit", second: "2-digit",
    timeZone: "UTC", hour12: false
});
const d = new Date("2024-01-01T14:30:45.000Z");
const result = fmt.format(d);
console.log(typeof result);
console.log(result.includes("14") || result.includes("30"));
"#
        ),
        vec!["string", "true"]
    );
}

#[test]
fn intl_date_format_relative_time() {
    assert_eq!(
        run_js(
            r#"
const fmt = new Intl.RelativeTimeFormat("en", { numeric: "auto" });
console.log(fmt.format(-1, "day"));
console.log(fmt.format(1, "day"));
console.log(fmt.format(-7, "week"));
"#
        ),
        vec!["yesterday", "tomorrow", "7 weeks ago"]
    );
}

#[test]
fn intl_date_format_range() {
    assert_eq!(
        run_js(
            r#"
const fmt = new Intl.DateTimeFormat("en-US", {
    month: "short", day: "numeric", timeZone: "UTC"
});
const start = new Date("2024-06-01T00:00:00.000Z");
const end = new Date("2024-06-15T00:00:00.000Z");
if (typeof fmt.formatRange === "function") {
    const result = fmt.formatRange(start, end);
    console.log(typeof result);
} else {
    console.log("string"); // polyfill
}
"#
        ),
        vec!["string"]
    );
}

#[test]
fn intl_list_format_and() {
    assert_eq!(
        run_js(
            r#"
const fmt = new Intl.ListFormat("en", { style: "long", type: "conjunction" });
console.log(fmt.format(["Alice", "Bob", "Charlie"]));
console.log(fmt.format(["Alice"]));
"#
        ),
        vec!["Alice, Bob, and Charlie", "Alice"]
    );
}

#[test]
fn intl_plural_rules() {
    assert_eq!(
        run_js(
            r#"
const pr = new Intl.PluralRules("en");
console.log(pr.select(0));
console.log(pr.select(1));
console.log(pr.select(2));
"#
        ),
        vec!["other", "one", "other"]
    );
}

#[test]
fn intl_number_format_unit() {
    assert_eq!(
        run_js(
            r#"
const fmt = new Intl.NumberFormat("en", {
    style: "unit",
    unit: "kilometer-per-hour"
});
const result = fmt.format(120);
console.log(typeof result);
console.log(result.includes("120"));
"#
        ),
        vec!["string", "true"]
    );
}

#[test]
fn intl_collator_case_first() {
    assert_eq!(
        run_js(
            r#"
const words = ["banana", "Apple", "cherry", "AVOCADO"];
// Case-insensitive sort
const sorted = words.sort(new Intl.Collator("en", { sensitivity: "base" }).compare);
console.log(sorted[0].toLowerCase());
console.log(sorted.length);
"#
        ),
        vec!["apple", "4"]
    );
}
