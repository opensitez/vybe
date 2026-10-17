//! Worker ownership and identity gates for committed frontend receipts.
use super::*;
use crate::frontend_preparation::{self as frontend, Capabilities, FrontendPreparation, PreparationSink};
use std::path::Path;
use std::sync::{Condvar, Mutex};
use std::sync::atomic::AtomicUsize;

fn module(source: &str) -> Module {
    Module { name: source.into(), language: vybe_ast::Lang::Unknown, body: vec![],
        imports: vec![], directives: Default::default(), canon: Default::default() }
}
fn language(parse: vybe_runtime::registry::ParseFn) -> Language {
    Language { name: "carrier-worker-test", parse, profile_source: || "",
        emit_dispatch: None, normalize_class: None, register_tree: None, expand_source: None }
}
struct Hints(&'static str);
impl DeferredPreparationHints for Hints {
    fn take_ready_paths(&mut self, _: &mut dyn FnMut(&str) -> Option<String>) -> Vec<String> {
        if self.0.is_empty() { vec![] } else { vec![std::mem::take(&mut self.0).into()] }
    }
    fn is_empty(&self) -> bool { self.0.is_empty() }
}
fn fixture(name: &str) -> PathBuf {
    let directory = std::env::temp_dir().join(format!("vybe-carrier-{}", uuid::Uuid::new_v4()));
    std::fs::create_dir_all(&directory).unwrap();
    let path = directory.join(name);
    std::fs::write(&path, "module").unwrap();
    path
}
fn scout(_: &str, _: &Path) -> Box<dyn DeferredPreparationHints> { Box::new(Hints("scout")) }

#[test]
fn committed_worker_receipts_drain_both_messages_without_second_parse() {
    static GATE: Mutex<bool> = Mutex::new(false);
    static SIGNAL: Condvar = Condvar::new();
    static CALLS: AtomicUsize = AtomicUsize::new(0);
    fn parser(_: &str) -> Result<Module,String> { panic!("provider owns parsing") }
    fn prepare(source: &str, _: &Path, sink: &mut PreparationSink<'_>) -> Result<FrontendPreparation,String> {
        CALLS.fetch_add(1, Ordering::SeqCst);
        sink(Box::new(Hints("committed")));
        let (release, timeout) = SIGNAL.wait_timeout_while(GATE.lock().unwrap(),
            std::time::Duration::from_secs(10), |release| !*release).unwrap();
        assert!(!timeout.timed_out(), "test must release the worker after commitment");
        drop(release);
        Ok(FrontendPreparation { module: module(source), hints: None })
    }
    frontend::register(parser, Capabilities { prepare: Some(prepare), early_hints: true, ..Default::default() });
    let path = fixture("drain.input");
    let selected = language(parser);
    let pool = PreparationPool::new(1, 4).unwrap();
    let ticket = pool.submit_with_hints(path.clone(), selected, Some(scout)).unwrap();
    let mut seen = vec![];
    for _ in 0..2 {
        let (_, mut hints) = ticket.hint_receiver().unwrap().recv_timeout(std::time::Duration::from_secs(5)).unwrap();
        seen.extend(hints.take_ready_paths(&mut |_| None));
    }
    assert_eq!(seen, ["scout", "committed"]);
    assert!(!ticket.is_ready());
    *GATE.lock().unwrap() = true; SIGNAL.notify_all();
    let prepared = ticket.wait_with_early_hints(|_| panic!("messages already consumed")).unwrap();
    assert!(prepared.preparation_hints_complete);
    assert!(prepared.matches_frontend(selected));
    assert_eq!(prepared.module.unwrap().name, "module");
    assert_eq!(CALLS.load(Ordering::SeqCst), 1);
    std::fs::remove_dir_all(path.parent().unwrap()).unwrap();
}

#[test]
fn dropped_parent_cancels_lowering_after_committed_receipt() {
    static GATE: Mutex<bool> = Mutex::new(false);
    static SIGNAL: Condvar = Condvar::new();
    fn parser(source: &str) -> Result<Module,String> { Ok(module(source)) }
    fn prepare(source: &str, _: &Path, sink: &mut PreparationSink<'_>) -> Result<FrontendPreparation,String> {
        sink(Box::new(Hints("committed")));
        let (release, timeout) = SIGNAL.wait_timeout_while(GATE.lock().unwrap(),
            std::time::Duration::from_secs(10), |release| !*release).unwrap();
        assert!(!timeout.timed_out()); drop(release);
        Ok(FrontendPreparation { module: module(source), hints: None })
    }
    frontend::register(parser, Capabilities { prepare: Some(prepare), ..Default::default() });
    let path = fixture("cancel.input");
    let pool = PreparationPool::new(1, 4).unwrap();
    let calls = Arc::new(AtomicUsize::new(0));
    let post = calls.clone();
    let ticket = pool.submit_with_post_prepare(path.clone(), language(parser), None,
        Some(Box::new(move |_| { post.fetch_add(1, Ordering::SeqCst); }))).unwrap();
    ticket.hint_receiver().unwrap().recv_timeout(std::time::Duration::from_secs(5)).unwrap();
    drop(ticket);
    *GATE.lock().unwrap() = true; SIGNAL.notify_all();
    frontend::register(parser, Capabilities::default());
    // Same one-worker queue provides an explicit completion barrier.
    assert!(pool.submit(path.clone(), language(parser)).unwrap().wait().unwrap().module.is_ok());
    assert_eq!(calls.load(Ordering::SeqCst), 0);
    std::fs::remove_dir_all(path.parent().unwrap()).unwrap();
}

#[test]
fn prepared_identity_rejects_resolver_provider_and_profile_drift() {
    static PROFILE: AtomicBool = AtomicBool::new(false);
    fn parser(source: &str) -> Result<Module,String> { Ok(module(source)) }
    fn profile() -> &'static str { if PROFILE.load(Ordering::SeqCst) { "changed" } else { "original" } }
    fn prepare(source: &str, _: &Path, _: &mut PreparationSink<'_>) -> Result<FrontendPreparation,String> {
        Ok(FrontendPreparation { module: module(source), hints: Some(Box::new(Hints(""))) })
    }
    fn resolver(_: &Path, _: &Path, target: &str, _: bool) -> PathBuf { target.into() }
    let mut selected = language(parser); selected.profile_source = profile;
    let path = fixture("identity.input");
    let pool = PreparationPool::new(1, 4).unwrap();
    let prepared = pool.submit(path.clone(), selected).unwrap().wait().unwrap();
    assert!(prepared.matches_frontend(selected));
    PROFILE.store(true, Ordering::SeqCst);
    assert!(!prepared.matches_frontend(selected));
    PROFILE.store(false, Ordering::SeqCst);
    frontend::register(parser, Capabilities { resolve_path: Some(resolver), ..Default::default() });
    assert!(!prepared.matches_frontend(selected));
    frontend::register(parser, Capabilities { prepare: Some(prepare), ..Default::default() });
    assert!(!prepared.matches_frontend(selected));
    frontend::register(parser, Capabilities::default());
    assert!(prepared.matches_frontend(selected));
    assert!(!prepared.matches_source(&path, "edited"));
    std::fs::remove_dir_all(path.parent().unwrap()).unwrap();
}

#[test]
fn completed_worker_wait_drains_scout_and_commit_once() {
    fn parser(_: &str) -> Result<Module,String> { panic!("provider owns parsing") }
    fn prepare(source: &str, _: &Path, sink: &mut PreparationSink<'_>) -> Result<FrontendPreparation,String> {
        sink(Box::new(Hints("committed")));
        Ok(FrontendPreparation { module: module(source), hints: None })
    }
    frontend::register(parser, Capabilities { prepare: Some(prepare), early_hints: true, ..Default::default() });
    let path = fixture("queued-drain.input");
    let pool = PreparationPool::new(1, 4).unwrap();
    let ticket = pool.submit_with_hints(path.clone(), language(parser), Some(scout)).unwrap();
    let mut seen = vec![];
    let prepared = ticket.wait_with_early_hints(|(_, mut hints)| seen.extend(hints.take_ready_paths(&mut |_| None))).unwrap();
    assert_eq!(seen, ["scout", "committed"]);
    assert!(prepared.preparation_hints_complete);
    assert_eq!(prepared.module.unwrap().name, "module");
    std::fs::remove_dir_all(path.parent().unwrap()).unwrap();
}
