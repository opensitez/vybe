//! Painting a document into a host pixmap, as webcore's `renderer` exposes it.
//!
//! The system browser composites the page into its own window, out of this
//! process, so there is no frame here to paint from. `render` leaves the pixmap
//! untouched and says so once, rather than handing back a blank frame as if it
//! were the page.

use std::sync::atomic::{AtomicBool, Ordering};

use tiny_skia::Pixmap;

use crate::layout::LayoutEngine;
use crate::types::Document;

#[derive(Debug, Default)]
pub struct Renderer {
    layout: LayoutEngine,
    scale: f32,
}

impl Renderer {
    pub fn new() -> Self {
        Self { layout: LayoutEngine::default(), scale: 1.0 }
    }

    pub fn set_scale(&mut self, scale: f32) {
        self.scale = scale;
    }

    pub fn layout_engine(&mut self) -> &mut LayoutEngine {
        &mut self.layout
    }

    pub fn render(&mut self, _doc: &mut Document, _pixmap: &mut Pixmap, _scale: f32) {
        static SAID: AtomicBool = AtomicBool::new(false);
        if !SAID.swap(true, Ordering::AcqRel) {
            eprintln!("osbrowser: the system browser paints its own window; there is no in-process frame to capture");
        }
    }
}
