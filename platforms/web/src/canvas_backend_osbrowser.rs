//! Browser-native CanvasRenderingContext2D behind `web:canvas`.

use std::sync::Arc;

use serde::Deserialize;
use serde_json::{Value, json};

use crate::canvas_backend::{
    self, CanvasBackend, Op2D, Query2D, Query2DValue, TextMetrics2D
};

struct OsBrowserCanvas;

fn target_parts(target: &str) -> (u64, String) {
    if let Some(rest) = target.strip_prefix('d') {
        if let Some((document, node)) = rest.split_once(':') {
            if let Ok(document) = document.parse() {
                return (document, node.to_string());
            }
        }
    }
    (crate::html::active_document(), target.to_string())
}

#[derive(Deserialize)]
enum RemoteQuery {
    Absent,
    Bool(bool),
    Text(String),
    Matrix([f32; 6]),
    Floats(Vec<f32>),
    Bytes(Vec<u8>),
    Pixels { data: Vec<u8>, width: u32, height: u32 ,},
    SourceImage { data: Vec<u8>, width: u32, height: u32, origin_clean: bool ,},
    Metrics(TextMetrics2D),
    ContextAttributes {
        alpha: bool,
        desynchronized: bool,
        color_space: String,
        color_type: String,
        will_read_frequently: bool,
    },
}

impl From<RemoteQuery> for Query2DValue {
    fn from(value: RemoteQuery) -> Self {
        match value {
            RemoteQuery::Absent => Self::Absent,
            RemoteQuery::Bool(value) => Self::Bool(value),
            RemoteQuery::Text(value) => Self::Text(value),
            RemoteQuery::Matrix(value) => Self::Matrix(value),
            RemoteQuery::Floats(value) => Self::Floats(value),
            RemoteQuery::Bytes(value) => Self::Bytes(value),
            RemoteQuery::Pixels {
                data,
                width,
                height,
            } => Self::Pixels {
                data,
                width,
                height,
            },
            RemoteQuery::SourceImage {
                data,
                width,
                height,
                origin_clean,
            } => Self::SourceImage {
                data,
                width,
                height,
                origin_clean,
            },
            RemoteQuery::Metrics(value) => Self::Metrics(value),
            RemoteQuery::ContextAttributes {
                alpha, desynchronized, color_space, color_type, will_read_frequently,
            } => Self::ContextAttributes {
                alpha, desynchronized, color_space, color_type, will_read_frequently,
            },
        }
    }
}

fn send(target: &str, operation: &str, payload: Value) -> Result<Value, String> {
    let (document, target) = target_parts(target);
    crate::engine_osbrowser::send(document, operation, json!({
        "target": target,
        "payload": payload,
    }),)
}

impl CanvasBackend for OsBrowserCanvas {
    fn ensure(&self, target: &str) {
        let _ = send(target, "CanvasEnsure", Value::Null);
    }

    fn apply(&self, target: &str, op: Op2D) {
        if let Ok(op) = serde_json::to_value(op) {
            if let Err(error) = send(target, "CanvasApply", op) {
                eprintln!("osbrowser canvas: {error}");
            }
        }
    }

    fn query(&self, target: &str, query: Query2D) -> Query2DValue {
        let Ok(query) = serde_json::to_value(query) else {
            return Query2DValue::Absent;
        };
        send(target, "CanvasQuery", query)
            .ok()
            .and_then(|value| serde_json::from_value::<RemoteQuery>(value).ok())
            .map(Into::into)
            .unwrap_or(Query2DValue::Absent)
    }

    fn clear_all(&self, target: &str) {
        let _ = send(target, "CanvasClearAll", Value::Null);
    }
}

pub fn install() {
    canvas_backend::set_backend(Arc::new(OsBrowserCanvas));
}
