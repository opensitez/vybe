//! CSS filter functions, for the canvas `filter` attribute.
//!
//! HTML §4.12.5 gives `CanvasRenderingContext2D.filter` the same grammar as the
//! CSS `filter` property, so a canvas has to understand `blur(4px)`,
//! `drop-shadow(2px 2px 4px black)` and the seven colour-matrix functions
//! whether or not the toolkit around it lays out CSS. This module is that
//! understanding, and the canvas is its only caller.
//!
//! The colour maths is [Filter Effects Module Level
//! 1](https://drafts.fxtf.org/filter-effects/#ShorthandEquivalents), which
//! defines each shorthand as an SVG filter primitive — that is where the
//! luminance coefficients and the hue-rotation matrix below come from, rather
//! than from any one implementation.

use tiny_skia::{Pixmap, PremultipliedColorU8};

use super::Color;

/// A parsed CSS `filter` value: the functions, in the order they were written.
#[derive(Clone, Debug, Default)]
pub struct CssFilters {
    pub ops: Vec<FilterOp>,
}

impl CssFilters {
    pub fn is_empty(&self) -> bool {
        self.ops.is_empty()
    }
}

/// One filter function.
///
/// The fractional arguments are already normalised out of percentages, so
/// `brightness(50%)` and `brightness(0.5)` arrive here identically — which is
/// what the grammar says they mean. Lengths are in CSS pixels and angles in
/// degrees.
#[derive(Clone, Copy, Debug)]
pub enum FilterOp {
    /// The argument is the standard deviation directly — unlike canvas
    /// `shadowBlur`, which names TWICE it. Keeping the two straight is the
    /// difference between a shadow that matches a browser and one that is
    /// twice as soft.
    Blur(f32),
    Brightness(f32),
    Contrast(f32),
    Grayscale(f32),
    HueRotate(f32),
    Invert(f32),
    Opacity(f32),
    Saturate(f32),
    Sepia(f32),
    DropShadow {
        dx: f32,
        dy: f32,
        blur: f32,
        color: Color,
    },
}

/// Parse a CSS `filter` value.
///
/// An unrecognised function is skipped rather than failing the list: the CSS
/// error-handling rule is that a declaration is dropped, and a canvas whose
/// `filter` silently kept working for the functions it does know is closer to
/// that than one that discards everything over a single unknown name.
pub fn parse_css_filter(value: &str) -> CssFilters {
    let mut ops = Vec::new();
    let trimmed = value.trim();
    if trimmed.is_empty() || trimmed.eq_ignore_ascii_case("none") {
        return CssFilters::default();
    }
    let mut rest = trimmed;
    while !rest.is_empty() {
        rest = rest.trim_start();
        if rest.is_empty() {
            break;
        }
        let Some(paren) = rest.find('(') else { break };
        let func = rest[..paren].trim().to_ascii_lowercase();
        let after = &rest[paren + 1..];
        let close = after.find(')').unwrap_or(after.len());
        let arg_str = after[..close].trim();
        rest = if close + 1 < after.len() {
            &after[close + 1..]
        } else {
            ""
        };

        let raw: f32 = arg_str
            .trim_end_matches('%')
            .trim_end_matches("px")
            .trim_end_matches("deg")
            .parse()
            .unwrap_or(0.0);
        // A percentage means the same thing as the bare fraction.
        let unit = if arg_str.ends_with('%') {
            raw / 100.0
        } else {
            raw
        };

        match func.as_str() {
            "blur" => ops.push(FilterOp::Blur(raw)),
            "brightness" => ops.push(FilterOp::Brightness(unit)),
            "contrast" => ops.push(FilterOp::Contrast(unit)),
            "grayscale" => ops.push(FilterOp::Grayscale(unit)),
            "hue-rotate" => ops.push(FilterOp::HueRotate(raw)),
            "invert" => ops.push(FilterOp::Invert(unit)),
            "opacity" => ops.push(FilterOp::Opacity(unit)),
            "saturate" => ops.push(FilterOp::Saturate(unit)),
            "sepia" => ops.push(FilterOp::Sepia(unit)),
            "drop-shadow" => {
                let parts: Vec<&str> = arg_str.split_whitespace().collect();
                let len = |i: usize| {
                    parts
                        .get(i)
                        .and_then(|s| s.trim_end_matches("px").parse::<f32>().ok())
                        .unwrap_or(0.0)
                };
                let color = parts
                    .get(3)
                    .and_then(|s| crate::layout::parse_color(s))
                    .map(|(r, g, b, a)| Color { r, g, b, a })
                    .unwrap_or(Color::BLACK);
                ops.push(FilterOp::DropShadow {
                    dx: len(0),
                    dy: len(1),
                    blur: len(2),
                    color,
                });
            }
            _ => {}
        }
    }
    CssFilters { ops }
}

/// Apply one colour-matrix filter in place.
///
/// Blur and drop-shadow are NOT here: they move pixels between positions rather
/// than mapping each colour to another, so they need a whole second buffer and
/// live in `effects`. Everything below is per-pixel and in place.
///
/// Each pixel is un-premultiplied before the transform and re-premultiplied
/// after. Skipping that step tints every partially transparent pixel toward
/// black, because a premultiplied channel is already scaled by alpha and the
/// filter maths is defined on the unscaled colour.
pub fn apply_color_matrix(pixmap: &mut Pixmap, op: &FilterOp) {
    let map: Box<dyn Fn(f32, f32, f32) -> (f32, f32, f32)> = match op {
        // Handled elsewhere — they are not per-pixel maps.
        FilterOp::Blur(_) | FilterOp::DropShadow { .. } => return,
        FilterOp::Brightness(v) => {
            let v = *v;
            Box::new(move |r, g, b| ((r * v).min(255.0), (g * v).min(255.0), (b * v).min(255.0)))
        }
        FilterOp::Contrast(v) => {
            let v = *v;
            Box::new(move |r, g, b| {
                let adj = |c: f32| (((c / 255.0 - 0.5) * v + 0.5) * 255.0).clamp(0.0, 255.0);
                (adj(r), adj(g), adj(b))
            })
        }
        FilterOp::Grayscale(v) => {
            let v = *v;
            Box::new(move |r, g, b| {
                let lum = 0.2126 * r + 0.7152 * g + 0.0722 * b;
                let mix = |c: f32| c * (1.0 - v) + lum * v;
                (mix(r), mix(g), mix(b))
            })
        }
        FilterOp::HueRotate(deg) => {
            let rad = deg * std::f32::consts::PI / 180.0;
            let (cos, sin) = (rad.cos(), rad.sin());
            Box::new(move |r, g, b| {
                let (r, g, b) = (r / 255.0, g / 255.0, b / 255.0);
                let out = |m: [f32; 3]| (m[0] * r + m[1] * g + m[2] * b) * 255.0;
                (
                    out([
                        0.213 + 0.787 * cos - 0.213 * sin,
                        0.715 - 0.715 * cos - 0.715 * sin,
                        0.072 - 0.072 * cos + 0.928 * sin,
                    ]),
                    out([
                        0.213 - 0.213 * cos + 0.143 * sin,
                        0.715 + 0.285 * cos + 0.140 * sin,
                        0.072 - 0.072 * cos - 0.283 * sin,
                    ]),
                    out([
                        0.213 - 0.213 * cos - 0.787 * sin,
                        0.715 - 0.715 * cos + 0.715 * sin,
                        0.072 + 0.928 * cos + 0.072 * sin,
                    ]),
                )
            })
        }
        FilterOp::Invert(v) => {
            let v = *v;
            Box::new(move |r, g, b| {
                let inv = |c: f32| c * (1.0 - v) + (255.0 - c) * v;
                (inv(r), inv(g), inv(b))
            })
        }
        // `opacity()` scales ALPHA, which the colour map cannot reach — done
        // below instead of here.
        FilterOp::Opacity(v) => {
            scale_alpha(pixmap, *v);
            return;
        }
        FilterOp::Saturate(v) => {
            let v = *v;
            Box::new(move |r, g, b| {
                let lum = 0.213 * r + 0.715 * g + 0.072 * b;
                let mix = |c: f32| lum + (c - lum) * v;
                (mix(r), mix(g), mix(b))
            })
        }
        FilterOp::Sepia(v) => {
            let v = *v;
            Box::new(move |r, g, b| {
                let sr = 0.393 * r + 0.769 * g + 0.189 * b;
                let sg = 0.349 * r + 0.686 * g + 0.168 * b;
                let sb = 0.272 * r + 0.534 * g + 0.131 * b;
                (
                    r * (1.0 - v) + sr * v,
                    g * (1.0 - v) + sg * v,
                    b * (1.0 - v) + sb * v,
                )
            })
        }
    };

    for px in pixmap.pixels_mut() {
        let a = px.alpha();
        if a == 0 {
            continue;
        }
        let af = a as f32 / 255.0;
        let (r, g, b) = map(
            px.red() as f32 / af,
            px.green() as f32 / af,
            px.blue() as f32 / af,
        );
        let back = |c: f32| (c * af).round().clamp(0.0, 255.0) as u8;
        if let Some(p) = PremultipliedColorU8::from_rgba(back(r), back(g), back(b), a) {
            *px = p;
        }
    }
}

/// `opacity(v)` — scale every pixel's alpha.
///
/// The colour channels scale with it because they are premultiplied, which is
/// exactly the relationship premultiplication encodes: a pixel at half alpha
/// carries half its colour.
fn scale_alpha(pixmap: &mut Pixmap, v: f32) {
    let v = v.clamp(0.0, 1.0);
    for px in pixmap.pixels_mut() {
        let scale = |c: u8| (c as f32 * v).round().clamp(0.0, 255.0) as u8;
        if let Some(p) = PremultipliedColorU8::from_rgba(
            scale(px.red()),
            scale(px.green()),
            scale(px.blue()),
            scale(px.alpha()),
        ) {
            *px = p;
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_filter_list_keeps_the_order_it_was_written_in() {
        // `blur(2px) brightness(2)` is not `brightness(2) blur(2px)` — the
        // second brightens what the blur already averaged.
        let f = parse_css_filter("blur(2px) brightness(2)");
        assert_eq!(f.ops.len(), 2);
        assert!(matches!(f.ops[0], FilterOp::Blur(v) if (v - 2.0).abs() < 1e-6));
        assert!(matches!(f.ops[1], FilterOp::Brightness(v) if (v - 2.0).abs() < 1e-6));
    }

    #[test]
    fn a_percentage_means_the_same_as_the_fraction() {
        let pct = parse_css_filter("grayscale(50%)");
        let frac = parse_css_filter("grayscale(0.5)");
        let val = |f: &CssFilters| match f.ops[0] {
            FilterOp::Grayscale(v) => v,
            _ => panic!("expected grayscale"),
        };
        assert!((val(&pct) - val(&frac)).abs() < 1e-6);
    }

    #[test]
    fn none_and_empty_are_no_filter_at_all() {
        assert!(parse_css_filter("none").is_empty());
        assert!(parse_css_filter("  ").is_empty());
        assert!(parse_css_filter("NONE").is_empty());
    }

    #[test]
    fn an_unknown_function_does_not_discard_the_rest() {
        let f = parse_css_filter("sparkle(3) invert(1)");
        assert_eq!(f.ops.len(), 1);
        assert!(matches!(f.ops[0], FilterOp::Invert(_)));
    }

    #[test]
    fn drop_shadow_reads_its_lengths_and_its_colour() {
        let f = parse_css_filter("drop-shadow(2px 3px 4px #ff0000)");
        match f.ops[0] {
            FilterOp::DropShadow {
                dx,
                dy,
                blur,
                color,
            } => {
                assert_eq!((dx, dy, blur), (2.0, 3.0, 4.0));
                assert_eq!((color.r, color.g, color.b), (255, 0, 0));
            }
            _ => panic!("expected a drop shadow"),
        }
    }

    /// A pixmap of one colour at one alpha.
    fn flat(r: u8, g: u8, b: u8, a: u8) -> Pixmap {
        let mut p = Pixmap::new(4, 4).expect("a pixmap");
        for px in p.pixels_mut() {
            let pm = |c: u8| ((c as u32 * a as u32) / 255) as u8;
            *px = PremultipliedColorU8::from_rgba(pm(r), pm(g), pm(b), a).expect("a colour");
        }
        p
    }

    #[test]
    fn a_filter_on_a_translucent_pixel_does_not_darken_it() {
        // The un-premultiply step. Applied to premultiplied channels directly,
        // `grayscale(1)` on a half-transparent red comes out darker than the
        // same colour at full alpha, because the maths would be running on a
        // colour already scaled by alpha.
        let mut half = flat(255, 0, 0, 128);
        let mut full = flat(255, 0, 0, 255);
        apply_color_matrix(&mut half, &FilterOp::Grayscale(1.0));
        apply_color_matrix(&mut full, &FilterOp::Grayscale(1.0));

        let unpre = |p: &Pixmap| {
            let px = p.pixels()[0];
            let af = px.alpha() as f32 / 255.0;
            px.red() as f32 / af
        };
        assert!(
            (unpre(&half) - unpre(&full)).abs() < 3.0,
            "half-alpha grey {} differs from full-alpha grey {}",
            unpre(&half),
            unpre(&full)
        );
    }

    #[test]
    fn grayscale_makes_the_channels_equal() {
        let mut p = flat(255, 0, 0, 255);
        apply_color_matrix(&mut p, &FilterOp::Grayscale(1.0));
        let px = p.pixels()[0];
        assert_eq!(px.red(), px.green());
        assert_eq!(px.green(), px.blue());
    }

    #[test]
    fn invert_turns_black_into_white() {
        let mut p = flat(0, 0, 0, 255);
        apply_color_matrix(&mut p, &FilterOp::Invert(1.0));
        let px = p.pixels()[0];
        assert_eq!((px.red(), px.green(), px.blue()), (255, 255, 255));
    }

    #[test]
    fn opacity_scales_alpha_and_not_only_colour() {
        // `opacity()` is the one shorthand a colour matrix cannot express, so
        // it is the one that silently does nothing if it is routed through one.
        let mut p = flat(255, 255, 255, 255);
        apply_color_matrix(&mut p, &FilterOp::Opacity(0.5));
        let px = p.pixels()[0];
        assert!(
            (px.alpha() as i32 - 128).abs() <= 1,
            "alpha stayed at {}",
            px.alpha()
        );
    }
}
