//! The window the page is shown in, as webcore's `embedded_window` exposes it.
//!
//! Here that window is the system browser's: each call is asked of the page,
//! which answers for its own window.

use serde_json::json;

use crate::page;

fn visible_page() -> u64 {
    page::visible_page()
}

pub fn focus() {
    page::write(visible_page(), "window_focus", json!([]));
}

pub fn screen_size() -> Option<(f64, f64)> {
    page::read(visible_page(), "window_screen_size", json!([]))
}

pub fn resize_to(width: f64, height: f64) {
    page::write(visible_page(), "window_resize_to", json!([width, height]));
}

pub fn move_to(x: f64, y: f64) {
    page::write(visible_page(), "window_move_to", json!([x, y]));
}

pub fn screen_position() -> Option<(f64, f64)> {
    page::read(visible_page(), "window_screen_position", json!([]))
}
