//! Genuine registered node:http request state, not a replacement provider.
use std::sync::{Arc, Mutex};
use vybe_platform_node::http::{RequestBodyReader, RequestContext, ResponseState};
pub fn request() -> Arc<RequestContext> {
    Arc::new(RequestContext {
        method: "GET".into(), uri: "/probe".into(), path: "/probe".into(),
        query: String::new(), scheme: "http".into(), host: "localhost".into(),
        headers: Vec::new(), env: Default::default(), remote_addr: "127.0.0.1".into(),
        remote_port: 1234, body: Mutex::new(RequestBodyReader::empty()),
        response: Mutex::new(ResponseState::new(None)),
    })
}
