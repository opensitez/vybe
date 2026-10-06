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
    pub query: Arc<str>,
    pub params: Arc<[String]>,
}

pub(super) struct SqlState {
    pub conns: HashMap<u64, ConnEntry>,
    pub stmts: HashMap<u64, StmtEntry>,
}

pub(super) fn state() -> &'static Mutex<SqlState> {
    static S: OnceLock<Mutex<SqlState>> = OnceLock::new();
    S.get_or_init(|| {
        Mutex::new(SqlState {
            conns: HashMap::new(),
            stmts: HashMap::new(),
        })
    })
}

/// VM hot-reset (bucket C/D): drop all open connections and prepared statements
/// so a reused VM never carries a prior run's DB handles. See `vmhotresetplan.md`.
pub fn reset() {
    let old = if let Ok(mut s) = state().lock() {
        s.stmts.clear();
        std::mem::take(&mut s.conns)
    } else {
        return;
    };
    drop(old);
}
