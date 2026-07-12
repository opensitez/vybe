/// String processing — parsing, formatting, transformation utilities
use super::helpers::run_js;

#[test]
fn parse_query_string() {
    assert_eq!(
        run_js(
            r#"
function parseQS(qs) {
    return Object.fromEntries(
        qs.split("&").map(p => {
            const [k, v] = p.split("=");
            return [decodeURIComponent(k), decodeURIComponent(v ?? "")];
        })
    );
}
const q = parseQS("name=Alice&age=30&city=New%20York");
console.log(q.name);
console.log(q.age);
console.log(q.city);
"#
        ),
        vec!["Alice", "30", "New York"]
    );
}

#[test]
fn format_bytes() {
    assert_eq!(
        run_js(
            r#"
function formatBytes(bytes) {
    const units = ["B","KB","MB","GB","TB"];
    let i = 0;
    while (bytes >= 1024 && i < units.length - 1) { bytes /= 1024; i++; }
    return bytes.toFixed(i === 0 ? 0 : 2) + " " + units[i];
}
console.log(formatBytes(0));
console.log(formatBytes(1024));
console.log(formatBytes(1024 * 1024));
console.log(formatBytes(1500));
"#
        ),
        vec!["0 B", "1.00 KB", "1.00 MB", "1.46 KB"]
    );
}

#[test]
fn slug_generation() {
    assert_eq!(
        run_js(
            r#"
function slugify(text) {
    return text.toLowerCase()
        .replace(/[^\w\s-]/g, "")
        .replace(/[\s_]+/g, "-")
        .replace(/^-+|-+$/g, "");
}
console.log(slugify("Hello World!"));
console.log(slugify("  The Quick Brown Fox  "));
console.log(slugify("Hello---World"));
"#
        ),
        vec!["hello-world", "the-quick-brown-fox", "hello---world"]
    );
}

#[test]
fn levenshtein_distance() {
    assert_eq!(
        run_js(
            r#"
function levenshtein(a, b) {
    const m = a.length, n = b.length;
    const dp = Array.from({length: m+1}, (_, i) => Array.from({length: n+1}, (_, j) => i || j));
    for (let i = 1; i <= m; i++) {
        for (let j = 1; j <= n; j++) {
            if (a[i-1] === b[j-1]) dp[i][j] = dp[i-1][j-1];
            else dp[i][j] = 1 + Math.min(dp[i-1][j], dp[i][j-1], dp[i-1][j-1]);
        }
    }
    return dp[m][n];
}
console.log(levenshtein("kitten", "sitting"));
console.log(levenshtein("", "hello"));
console.log(levenshtein("abc", "abc"));
"#
        ),
        vec!["3", "5", "0"]
    );
}

#[test]
fn template_engine_simple() {
    assert_eq!(
        run_js(
            r#"
function render(template, data) {
    return template.replace(/\{\{(\w+)\}\}/g, (_, key) => data[key] ?? "");
}
const tmpl = "Hello, {{name}}! You are {{age}} years old.";
console.log(render(tmpl, { name: "Alice", age: 30 }));
console.log(render("{{missing}} world", {}));
"#
        ),
        vec!["Hello, Alice! You are 30 years old.", " world"]
    );
}

#[test]
fn string_tokenizer() {
    assert_eq!(
        run_js(
            r#"
function* tokenize(str) {
    const re = /(\d+\.?\d*)|([a-zA-Z_]\w*)|([+\-*\/()=])/g;
    let m;
    while ((m = re.exec(str)) !== null) {
        if (m[1]) yield { type: "number", value: m[1] };
        else if (m[2]) yield { type: "ident", value: m[2] };
        else yield { type: "op", value: m[3] };
    }
}
const tokens = [...tokenize("x = 3.14 + y")];
console.log(tokens.length);
console.log(tokens[0].type + ":" + tokens[0].value);
console.log(tokens[2].type + ":" + tokens[2].value);
"#
        ),
        vec!["5", "ident:x", "number:3.14"]
    );
}

#[test]
fn number_formatting_custom() {
    assert_eq!(
        run_js(
            r#"
function formatNum(n, decimals = 2, sep = ",") {
    const [int, dec] = n.toFixed(decimals).split(".");
    const formatted = int.replace(/\B(?=(\d{3})+(?!\d))/g, sep);
    return dec ? `${formatted}.${dec}` : formatted;
}
console.log(formatNum(1234567.89));
console.log(formatNum(1000, 0));
console.log(formatNum(42.1234, 3));
"#
        ),
        vec!["1,234,567.89", "1,000", "42.123"]
    );
}

#[test]
fn csv_serializer() {
    assert_eq!(
        run_js(
            r#"
function toCSV(data, headers) {
    const escape = v => /[,"\n]/.test(String(v)) ? `"${String(v).replace(/"/g, '""')}"` : String(v);
    const rows = data.map(row => headers.map(h => escape(row[h])).join(","));
    return [headers.join(","), ...rows].join("\n");
}
const data = [
    { name: "Alice", age: 30, city: "New York" },
    { name: "Bob", age: 25, city: "Los Angeles" },
];
const csv = toCSV(data, ["name", "age", "city"]);
const lines = csv.split("\n");
console.log(lines[0]);
console.log(lines[1]);
"#
        ),
        vec!["name,age,city", "Alice,30,New York"]
    );
}

#[test]
fn word_wrap() {
    assert_eq!(
        run_js(
            r#"
function wordWrap(text, width) {
    const words = text.split(" ");
    const lines = [];
    let line = "";
    for (const word of words) {
        if ((line + " " + word).trim().length <= width) {
            line = (line + " " + word).trim();
        } else {
            if (line) lines.push(line);
            line = word;
        }
    }
    if (line) lines.push(line);
    return lines;
}
const lines = wordWrap("The quick brown fox jumps over the lazy dog", 15);
console.log(lines[0]);
console.log(lines[1]);
"#
        ),
        vec!["The quick brown", "fox jumps over"]
    );
}

#[test]
fn diff_two_strings() {
    assert_eq!(
        run_js(
            r#"
function changedWords(a, b) {
    const wa = a.split(" "), wb = b.split(" ");
    const changes = [];
    const maxLen = Math.max(wa.length, wb.length);
    for (let i = 0; i < maxLen; i++) {
        if (wa[i] !== wb[i]) changes.push(i);
    }
    return changes;
}
console.log(changedWords("hello world foo", "hello bar foo").join(","));
console.log(changedWords("a b c", "a b c").length);
"#
        ),
        vec!["1", "0"]
    );
}
