use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Well-Formed Unicode Strings (`isWellFormed`, `toWellFormed`) (ES2024)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_string_is_well_formed_ascii_and_bmp() {
    let src = r#"
const str = "Hello World";
console.log(str.isWellFormed());
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_string_is_well_formed_paired_surrogates() {
    let src = r#"
const emoji = "😀🚀🎉";
console.log(emoji.isWellFormed());
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_string_is_well_formed_unpaired_lead_surrogate() {
    let src = r#"
const loneLead = "a\uD83Db";
console.log(loneLead.isWellFormed());
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_string_is_well_formed_unpaired_trail_surrogate() {
    let src = r#"
const loneTrail = "a\uDE00b";
console.log(loneTrail.isWellFormed());
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_string_to_well_formed_valid_string_unchanged() {
    let src = r#"
const str = "Valid 😀 String";
console.log(str.toWellFormed() === str);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_string_to_well_formed_replaces_unpaired_lead_surrogate() {
    let src = r#"
const loneLead = "a\uD83Db";
const wellFormed = loneLead.toWellFormed();
console.log(wellFormed + "|code=" + wellFormed.charCodeAt(1)); // Lone surrogate replaced by U+FFFD (65533 replacement character)!
"#;
    assert_eq!(run_js(src), vec!["ab|code=65533"]);
}

#[test]
fn test_js_string_to_well_formed_replaces_unpaired_trail_surrogate() {
    let src = r#"
const loneTrail = "a\uDE00b";
const wellFormed = loneTrail.toWellFormed();
console.log(wellFormed + "|code=" + wellFormed.charCodeAt(1));
"#;
    assert_eq!(run_js(src), vec!["ab|code=65533"]);
}

#[test]
fn test_js_string_to_well_formed_replaces_reversed_surrogate_pair() {
    let src = r#"
const reversedSurrogates = "\uDE00\uD83D"; // Trail surrogate followed by lead surrogate
const wellFormed = reversedSurrogates.toWellFormed();
console.log(wellFormed.length + "|isWellFormed=" + wellFormed.isWellFormed());
"#;
    assert_eq!(run_js(src), vec!["2|isWellFormed=true"]);
}

#[test]
fn test_js_string_is_well_formed_empty_string() {
    let src = r#"
console.log("".isWellFormed());
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_string_to_well_formed_empty_string() {
    let src = r#"
console.log("".toWellFormed() === "");
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_string_is_well_formed_property_descriptor() {
    let src = r#"
const desc = Object.getOwnPropertyDescriptor(String.prototype, "isWellFormed");
console.log(`${desc.writable}:${desc.enumerable}:${desc.configurable}:${String.prototype.isWellFormed.length}`);
"#;
    assert_eq!(run_js(src), vec!["true:false:true:0"]);
}

#[test]
fn test_js_string_to_well_formed_property_descriptor() {
    let src = r#"
const desc = Object.getOwnPropertyDescriptor(String.prototype, "toWellFormed");
console.log(`${desc.writable}:${desc.enumerable}:${desc.configurable}:${String.prototype.toWellFormed.length}`);
"#;
    assert_eq!(run_js(src), vec!["true:false:true:0"]);
}

#[test]
fn test_js_string_is_well_formed_coerces_this_to_string() {
    let src = r#"
console.log(String.prototype.isWellFormed.call(12345));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_string_to_well_formed_coerces_this_to_string() {
    let src = r#"
const res = String.prototype.toWellFormed.call(12345);
console.log(typeof res + "|" + res);
"#;
    assert_eq!(run_js(src), vec!["string|12345"]);
}

#[test]
fn test_js_string_is_well_formed_null_this_throws_typeerror() {
    let src = r#"
try {
    String.prototype.isWellFormed.call(null);
} catch (e) {
    console.log("isWellFormed Null This TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["isWellFormed Null This TypeError"]);
}

#[test]
fn test_js_string_to_well_formed_encode_uri_component_safety() {
    let src = r#"
const malformed = "a\uD83Db";
const safe = malformed.toWellFormed();
console.log(encodeURIComponent(safe)); // encodeURIComponent throws URIError on malformed strings, toWellFormed makes it safe!
"#;
    assert_eq!(run_js(src), vec!["a%EF%BF%BDb"]);
}

#[test]
fn test_js_string_is_well_formed_name_property() {
    let src = r#"
console.log(String.prototype.isWellFormed.name);
"#;
    assert_eq!(run_js(src), vec!["isWellFormed"]);
}

#[test]
fn test_js_string_to_well_formed_name_property() {
    let src = r#"
console.log(String.prototype.toWellFormed.name);
"#;
    assert_eq!(run_js(src), vec!["toWellFormed"]);
}

#[test]
fn test_js_string_is_well_formed_multiple_unpaired_surrogates() {
    let src = r#"
const multipleLone = "\uD800\uD800\uD800";
console.log(multipleLone.isWellFormed());
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_string_to_well_formed_multiple_unpaired_surrogates_replacement() {
    let src = r#"
const multipleLone = "\uD800\uD800\uD800";
const fixed = multipleLone.toWellFormed();
console.log(fixed + "|len=" + fixed.length);
"#;
    assert_eq!(run_js(src), vec!["|len=3"]);
}

#[test]
fn test_js_string_to_well_formed_undefined_this_throws_typeerror() {
    let src = r#"
try {
    String.prototype.toWellFormed.call(undefined);
} catch (e) {
    console.log(e.name);
}
"#;
    assert_eq!(run_js(src), vec!["TypeError"]);
}

