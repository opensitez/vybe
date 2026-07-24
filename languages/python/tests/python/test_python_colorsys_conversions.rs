use super::helpers::run_python;

// ════════════════════════════════════════════════════════════
// Category: colorsys — color space conversions
// ════════════════════════════════════════════════════════════

#[test]
fn test_colorsys_rgb_to_hls_red() {
    let out = run_python(r#"
import colorsys
h, l, s = colorsys.rgb_to_hls(1.0, 0.0, 0.0)
print(round(h, 4))
print(round(l, 4))
"#);
    assert_eq!(out, vec!["0.0", "0.5"]);
}

#[test]
fn test_colorsys_hls_to_rgb_red() {
    let out = run_python(r#"
import colorsys
r, g, b = colorsys.hls_to_rgb(0.0, 0.5, 1.0)
print(round(r, 4))
print(round(g, 4))
print(round(b, 4))
"#);
    assert_eq!(out, vec!["1.0", "0.0", "0.0"]);
}

#[test]
fn test_colorsys_rgb_to_hsv_green() {
    let out = run_python(r#"
import colorsys
h, s, v = colorsys.rgb_to_hsv(0.0, 1.0, 0.0)
print(round(h, 4))
print(round(s, 4))
print(round(v, 4))
"#);
    assert_eq!(out, vec!["0.3333", "1.0", "1.0"]);
}

#[test]
fn test_colorsys_hsv_to_rgb_green() {
    let out = run_python(r#"
import colorsys
r, g, b = colorsys.hsv_to_rgb(1/3, 1.0, 1.0)
print(round(r, 4))
print(round(g, 4))
print(round(b, 4))
"#);
    assert_eq!(out, vec!["0.0", "1.0", "0.0"]);
}

#[test]
fn test_colorsys_rgb_to_yiq_white() {
    let out = run_python(r#"
import colorsys
y, i, q = colorsys.rgb_to_yiq(1.0, 1.0, 1.0)
print(round(y, 4))
print(round(i, 4))
print(round(q, 4))
"#);
    assert_eq!(out, vec!["1.0", "0.0", "0.0"]);
}

#[test]
fn test_colorsys_yiq_to_rgb_roundtrip() {
    let out = run_python(r#"
import colorsys
r0, g0, b0 = 0.2, 0.5, 0.8
y, i, q = colorsys.rgb_to_yiq(r0, g0, b0)
r1, g1, b1 = colorsys.yiq_to_rgb(y, i, q)
print(round(r1, 6) == round(r0, 6))
print(round(g1, 6) == round(g0, 6))
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_colorsys_hls_roundtrip() {
    let out = run_python(r#"
import colorsys
r0, g0, b0 = 0.3, 0.6, 0.9
h, l, s = colorsys.rgb_to_hls(r0, g0, b0)
r1, g1, b1 = colorsys.hls_to_rgb(h, l, s)
print(round(r1, 5) == round(r0, 5))
print(round(b1, 5) == round(b0, 5))
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_colorsys_hsv_roundtrip() {
    let out = run_python(r#"
import colorsys
r0, g0, b0 = 0.1, 0.4, 0.7
h, s, v = colorsys.rgb_to_hsv(r0, g0, b0)
r1, g1, b1 = colorsys.hsv_to_rgb(h, s, v)
print(round(r1, 5) == round(r0, 5))
print(round(g1, 5) == round(g0, 5))
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_colorsys_black_rgb_to_hsv() {
    let out = run_python(r#"
import colorsys
h, s, v = colorsys.rgb_to_hsv(0.0, 0.0, 0.0)
print(h, s, v)
"#);
    assert_eq!(out, vec!["0.0 0.0 0.0"]);
}

#[test]
fn test_colorsys_white_rgb_to_hsv() {
    let out = run_python(r#"
import colorsys
h, s, v = colorsys.rgb_to_hsv(1.0, 1.0, 1.0)
print(h, s, v)
"#);
    assert_eq!(out, vec!["0.0 0.0 1.0"]);
}

#[test]
fn test_colorsys_blue_hls() {
    let out = run_python(r#"
import colorsys
h, l, s = colorsys.rgb_to_hls(0.0, 0.0, 1.0)
print(round(h, 4))
"#);
    assert_eq!(out, vec!["0.6667"]);
}

#[test]
fn test_colorsys_hsv_to_rgb_black() {
    let out = run_python(r#"
import colorsys
r, g, b = colorsys.hsv_to_rgb(0.0, 0.0, 0.0)
print(r, g, b)
"#);
    assert_eq!(out, vec!["0.0 0.0 0.0"]);
}

#[test]
fn test_colorsys_hsv_to_rgb_white() {
    let out = run_python(r#"
import colorsys
r, g, b = colorsys.hsv_to_rgb(0.0, 0.0, 1.0)
print(r, g, b)
"#);
    assert_eq!(out, vec!["1.0 1.0 1.0"]);
}

#[test]
fn test_colorsys_hls_to_rgb_white() {
    let out = run_python(r#"
import colorsys
r, g, b = colorsys.hls_to_rgb(0.0, 1.0, 0.0)
print(r, g, b)
"#);
    assert_eq!(out, vec!["1.0 1.0 1.0"]);
}

#[test]
fn test_colorsys_hls_to_rgb_black() {
    let out = run_python(r#"
import colorsys
r, g, b = colorsys.hls_to_rgb(0.0, 0.0, 0.0)
print(r, g, b)
"#);
    assert_eq!(out, vec!["0.0 0.0 0.0"]);
}

#[test]
fn test_colorsys_cyan_rgb_to_hsv() {
    let out = run_python(r#"
import colorsys
h, s, v = colorsys.rgb_to_hsv(0.0, 1.0, 1.0)
print(round(h, 4))
print(s, v)
"#);
    assert_eq!(out, vec!["0.5", "1.0 1.0"]);
}

#[test]
fn test_colorsys_yiq_to_rgb_black() {
    let out = run_python(r#"
import colorsys
r, g, b = colorsys.yiq_to_rgb(0.0, 0.0, 0.0)
print(r, g, b)
"#);
    assert_eq!(out, vec!["0.0 0.0 0.0"]);
}

#[test]
fn test_colorsys_magenta_rgb_to_hls() {
    let out = run_python(r#"
import colorsys
h, l, s = colorsys.rgb_to_hls(1.0, 0.0, 1.0)
print(round(h, 4))
print(l)
"#);
    assert_eq!(out, vec!["0.8333", "0.5"]);
}

#[test]
fn test_colorsys_yellow_rgb_to_hsv() {
    let out = run_python(r#"
import colorsys
h, s, v = colorsys.rgb_to_hsv(1.0, 1.0, 0.0)
print(round(h, 4))
print(s, v)
"#);
    assert_eq!(out, vec!["0.1667", "1.0 1.0"]);
}

#[test]
fn test_colorsys_grey_rgb_to_hls() {
    let out = run_python(r#"
import colorsys
h, l, s = colorsys.rgb_to_hls(0.5, 0.5, 0.5)
print(h)
print(l)
print(s)
"#);
    assert_eq!(out, vec!["0.0", "0.5", "0.0"]);
}
