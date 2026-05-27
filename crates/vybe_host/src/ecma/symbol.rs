//! ECMA-262 §20.4 / §27.5 — Symbol + well-known symbols.
//!
//!   §20.4.1 Symbol(description?) — produces a unique primitive symbol.
//!   §20.4.2 Symbol.{for, keyFor} — global symbol registry.
//!   §20.4.x Symbol.{iterator, asyncIterator, toPrimitive, hasInstance,
//!     toStringTag, isConcatSpreadable, unscopables, match, matchAll,
//!     replace, search, split, species} — well-known symbols.
//!
//! Symbol equality semantics in Vybe: `Value::Symbol(Arc<str>)` — two
//! symbols are `===` iff they're cloned from the same `Arc<str>` (per
//! the `PartialEq` impl in `vybe_bytecode/src/value.rs`). Each
//! `Symbol(desc)` call mints a fresh `Arc<str>` so identity is unique.
//! `Symbol.for(key)` interns through a process-global registry so
//! repeat lookups return the same `Arc`.

use std::collections::{HashMap, HashSet};
use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};

// Well-known symbols — created once at register time, exposed through
// the namespace registry (see `crate::namespaces::value`).
struct WellKnown {
    iterator: Arc<str>,
    async_iterator: Arc<str>,
    to_primitive: Arc<str>,
    has_instance: Arc<str>,
    to_string_tag: Arc<str>,
    is_concat_spreadable: Arc<str>,
    unscopables: Arc<str>,
    match_: Arc<str>,
    match_all: Arc<str>,
    replace: Arc<str>,
    search: Arc<str>,
    split: Arc<str>,
    species: Arc<str>,
    dispose: Arc<str>,
    async_dispose: Arc<str>,
}

static WELL_KNOWN: std::sync::OnceLock<WellKnown> = std::sync::OnceLock::new();

fn well_known() -> &'static WellKnown {
    WELL_KNOWN.get_or_init(|| WellKnown {
        iterator:             Arc::from("@@iterator"),
        async_iterator:       Arc::from("@@asyncIterator"),
        to_primitive:         Arc::from("@@toPrimitive"),
        has_instance:         Arc::from("@@hasInstance"),
        to_string_tag:        Arc::from("@@toStringTag"),
        is_concat_spreadable: Arc::from("@@isConcatSpreadable"),
        unscopables:          Arc::from("@@unscopables"),
        match_:               Arc::from("@@match"),
        match_all:            Arc::from("@@matchAll"),
        replace:              Arc::from("@@replace"),
        search:               Arc::from("@@search"),
        split:                Arc::from("@@split"),
        species:              Arc::from("@@species"),
        dispose:              Arc::from("@@dispose"),
        async_dispose:        Arc::from("@@asyncDispose"),
    })
}

// Process-global Symbol.for(...) registry. Per spec §20.4.2.2 each key
// maps to one canonical symbol shared across realms (in our case, the
// VM process).
static REGISTRY: std::sync::OnceLock<Mutex<HashMap<String, Arc<str>>>> = std::sync::OnceLock::new();
static NO_DESCRIPTION_SYMBOLS: std::sync::OnceLock<Mutex<HashSet<usize>>> = std::sync::OnceLock::new();

fn registry() -> &'static Mutex<HashMap<String, Arc<str>>> {
    REGISTRY.get_or_init(|| Mutex::new(HashMap::new()))
}

fn no_description_symbols() -> &'static Mutex<HashSet<usize>> {
    NO_DESCRIPTION_SYMBOLS.get_or_init(|| Mutex::new(HashSet::new()))
}

pub(crate) fn canonical_property_key(sym: &Arc<str>) -> String {
    match sym.as_ref() {
        "@@iterator" => "iterator".to_string(),
        "@@asyncIterator" => "asyncIterator".to_string(),
        "@@toPrimitive" => "toprimitive".to_string(),
        "@@hasInstance" => "hasinstance".to_string(),
        "@@toStringTag" => "tostringtag".to_string(),
        "@@isConcatSpreadable" => "isconcatspreadable".to_string(),
        "@@unscopables" => "unscopables".to_string(),
        "@@match" => "symbolmatch".to_string(),
        "@@matchAll" => "symbolmatchall".to_string(),
        "@@replace" => "symbolreplace".to_string(),
        "@@search" => "symbolsearch".to_string(),
        "@@split" => "symbolsplit".to_string(),
        "@@species" => "species".to_string(),
        "@@dispose" => "dispose".to_string(),
        "@@asyncDispose" => "asyncdispose".to_string(),
        _ => format!("Symbol({})", sym),
    }
}

pub(crate) fn has_description(sym: &Arc<str>) -> bool {
    !no_description_symbols()
        .lock()
        .unwrap()
    .contains(&(sym.as_ptr() as usize))
}

pub fn register(vm: &mut VM) {
    // `Symbol(description?)` — fresh unique symbol. Description is
    // stored in the Arc<str> contents so `toString()` round-trips.
    vm.register_host_fn("ecma:symbol", "Symbol", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        // Spec §20.4.1.1: store description verbatim. toString wraps it
        // as `Symbol(<desc>)` per §20.4.3.3 — keeping the wrap there
        // means Display and toString agree on a single representation.
        let has_description = args.first().is_some_and(|v| !matches!(v, Value::Undefined));
        let desc = args.first()
            .filter(|v| !matches!(v, Value::Undefined))
            .map(|v| format!("{}", v))
            .unwrap_or_default();
        let symbol = Arc::<str>::from(desc.as_str());
        if !has_description {
            no_description_symbols()
                .lock()
                .unwrap()
                .insert(symbol.as_ptr() as usize);
        }
        Value::Symbol(symbol)
    }));

    // Symbol.for(key) — global registry lookup; creates if absent.
    vm.register_host_fn("ecma:symbol", "for", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let key = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        let mut reg = registry().lock().unwrap();
        let arc = reg.entry(key.clone()).or_insert_with(|| Arc::from(key.as_str())).clone();
        Value::Symbol(arc)
    }));

    // Symbol.keyFor(sym) — reverse lookup; returns string or undefined.
    vm.register_host_fn("ecma:symbol", "keyFor", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        if let Some(Value::Symbol(sym)) = args.first() {
            let reg = registry().lock().unwrap();
            for (key, val) in reg.iter() {
                if Arc::ptr_eq(sym, val) {
                    return Value::String(Arc::from(key.as_str()));
                }
            }
        }
        Value::Undefined
    }));

    // Well-known symbols exposed as 0-arg getters so they slot into
    // the same flat host_registry as `Number.MAX_SAFE_INTEGER` etc.
    let wk = well_known();
    register_constant(vm, "iterator",            wk.iterator.clone());
    register_constant(vm, "asyncIterator",       wk.async_iterator.clone());
    register_constant(vm, "toPrimitive",         wk.to_primitive.clone());
    register_constant(vm, "hasInstance",         wk.has_instance.clone());
    register_constant(vm, "toStringTag",         wk.to_string_tag.clone());
    register_constant(vm, "isConcatSpreadable",  wk.is_concat_spreadable.clone());
    register_constant(vm, "unscopables",         wk.unscopables.clone());
    register_constant(vm, "match",               wk.match_.clone());
    register_constant(vm, "matchAll",            wk.match_all.clone());
    register_constant(vm, "replace",             wk.replace.clone());
    register_constant(vm, "search",              wk.search.clone());
    register_constant(vm, "split",               wk.split.clone());
    register_constant(vm, "species",             wk.species.clone());
    register_constant(vm, "dispose",             wk.dispose.clone());
    register_constant(vm, "asyncDispose",        wk.async_dispose.clone());
}

fn register_constant(vm: &mut VM, name: &'static str, sym: Arc<str>) {
    vm.register_host_fn("ecma:symbol", name, Box::new(move |_ctx: &mut HostContext, _args: &[Value]| {
        Value::Symbol(sym.clone())
    }));
}
