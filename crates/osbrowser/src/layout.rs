//! Layout, as webcore's `layout` exposes it. The system browser lays the page
//! out itself, whenever a question about geometry needs it.

use crate::types::Document;

#[derive(Debug, Default)]
pub struct LayoutEngine {
    pub viewport_h: f32,
}

impl LayoutEngine {
    pub fn layout(&mut self, _doc: &mut Document, _viewport_width: f32) {}
}
