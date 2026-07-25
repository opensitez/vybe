use super::driver::SqlDriver;
use std::collections::HashMap;
use std::sync::{
    Arc, Mutex, OnceLock,
    atomic::{AtomicU64, Ordering},
};

static NEXT_ID: AtomicU64 = AtomicU64::new(1);
pub(super) fn next_id() -> u64 {
    NEXT_ID.fetch_add(1, Ordering::Relaxed)
}

pub(super) struct ConnEntry {
    pub driver: Arc<dyn SqlDriver>,
    pub url: String,
}

pub(super) struct StmtEntry {
    pub query: String,
    pub params: Vec<String>,
}

pub(super) struct SqlState {
    pub conns: HashMap<u64, ConnEntry>,
    pub stmts: HashMap<u64, StmtEntry>,
}

pub(super) fn state() -> Arc<Mutex<SqlState>> {
    static S: OnceLock<Arc<Mutex<SqlState>>> = OnceLock::new();
    S.get_or_init(|| {
        Arc::new(Mutex::new(SqlState {
            conns: HashMap::new(),
            stmts: HashMap::new(),
        }))
    })
    .clone()
}

/// VM hot-reset (bucket C/D): drop all open connections and prepared statements
/// so a reused VM never carries a prior run's DB handles. See `vmhotresetplan.md`.
pub fn reset() {
    if let Ok(mut s) = state().lock() {
        s.conns.clear();
        s.stmts.clear();
    }
}
