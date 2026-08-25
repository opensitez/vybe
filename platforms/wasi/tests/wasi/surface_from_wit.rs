//! The declared surface, READ FROM THE VENDORED WIT — not from a hand-written list.
//!
//! `interface_coverage.rs` asserts `declared ⊆ registered`, which is the right
//! assertion, but its "declared" side is a list maintained by hand. A gate is
//! only ever as complete as its list, and a list cannot report what was never
//! added to it: a function nobody typed in is invisible to it, and reads as
//! full coverage.
//!
//! That is not hypothetical. `wasi:http/client.send` was registered, gated
//! green, and returned "HTTPS not supported (use http://)" — the list said
//! everything declared was present, and it was, because the list did not know
//! what the WIT declared beyond it.
//!
//! So this module parses `proposals/WASI/proposals/*/wit/*.wit` at test time
//! and derives the surface from the spec itself. The list in
//! `interface_coverage.rs` is then checked AGAINST that, which turns "someone
//! remembered to add it" into "the WIT says so".
//!
//! ## What the parser handles, and what it deliberately does not
//!
//! Free functions, resource methods (`[method]r.f`), statics (`[static]r.f`)
//! and constructors (`[constructor]r`) — the four shapes a host registers. It
//! does NOT resolve `use` re-exports or `include`d worlds: a name is attributed
//! to the interface whose braces contain it. That matches how the host
//! registers, which is what is being compared.
//!
//! `run` is excluded: `command.wit` says `export run`, so it is a GUEST export,
//! not a host import, and no runtime registers it.

use std::collections::{BTreeMap, BTreeSet};
use std::path::{Path, PathBuf};

/// The six packages `@0.3.1` comprises.
const PACKAGES: &[&str] = &[
    "cli",
    "clocks",
    "filesystem",
    "http",
    "random",
    "sockets",
];

fn proposals_root() -> PathBuf {
    // CARGO_MANIFEST_DIR is `platforms/wasi`; the proposals are vendored at the
    // repository root, which is two levels up.
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../proposals/WASI/proposals")
        .canonicalize()
        .expect("vendored WASI proposals are missing — they are the spec this suite checks against")
}

/// Strip `//` comments so a doc line mentioning `func` is not read as one.
fn strip_comments(src: &str) -> String {
    src.lines()
        .map(|line| match line.find("//") {
            Some(at) => &line[..at],
            None => line,
        })
        .collect::<Vec<_>>()
        .join("\n")
}

/// The body of the brace-delimited block whose opening brace is at `open`.
fn block_at(src: &str, open: usize) -> (usize, usize) {
    let bytes = src.as_bytes();
    let mut depth = 0usize;
    let mut i = open;
    while i < bytes.len() {
        match bytes[i] {
            b'{' => depth += 1,
            b'}' => {
                depth -= 1;
                if depth == 0 {
                    return (open + 1, i);
                }
            }
            _ => {}
        }
        i += 1;
    }
    (open + 1, src.len())
}

/// Every `name: func`, `name: static func`, `name: async func` at this level.
fn free_functions(body: &str) -> Vec<(String, bool)> {
    let mut out = Vec::new();
    for line in body.lines() {
        let line = line.trim();
        let Some(colon) = line.find(':') else { continue };
        let name = line[..colon].trim();
        if name.is_empty() || !name.chars().all(|c| c.is_ascii_lowercase() || c.is_ascii_digit() || c == '-') {
            continue;
        }
        let rest = line[colon + 1..].trim();
        let is_static = rest.starts_with("static ");
        let rest = rest.strip_prefix("static ").unwrap_or(rest);
        let rest = rest.strip_prefix("async ").unwrap_or(rest);
        if rest.starts_with("func") {
            out.push((name.to_string(), is_static));
        }
    }
    out
}

/// `interface name -> the host-import spellings it declares`.
pub fn declared_surface(package: &str) -> BTreeMap<String, BTreeSet<String>> {
    let dir = proposals_root().join(package).join("wit");
    let mut out: BTreeMap<String, BTreeSet<String>> = BTreeMap::new();
    let entries = std::fs::read_dir(&dir)
        .unwrap_or_else(|e| panic!("cannot read {}: {e}", dir.display()));
    for entry in entries.flatten() {
        let path = entry.path();
        if path.extension().and_then(|e| e.to_str()) != Some("wit") {
            continue;
        }
        let src = strip_comments(&std::fs::read_to_string(&path).expect("readable wit"));
        let mut cursor = 0usize;
        while let Some(at) = src[cursor..].find("interface ") {
            let start = cursor + at + "interface ".len();
            let name_end = src[start..]
                .find(|c: char| c.is_whitespace() || c == '{')
                .map(|i| start + i)
                .unwrap_or(src.len());
            let iface = src[start..name_end].trim().to_string();
            let Some(brace_rel) = src[name_end..].find('{') else { break };
            let (b0, b1) = block_at(&src, name_end + brace_rel);
            let mut body = src[b0..b1].to_string();
            cursor = b1;

            let names = out.entry(iface).or_default();

            // Resources first, then blank them so their methods are not
            // re-read as free functions of the interface.
            let mut scan = 0usize;
            while let Some(rat) = body[scan..].find("resource ") {
                let rstart = scan + rat + "resource ".len();
                let rname_end = body[rstart..]
                    .find(|c: char| c.is_whitespace() || c == '{')
                    .map(|i| rstart + i)
                    .unwrap_or(body.len());
                let rname = body[rstart..rname_end].trim().to_string();
                let Some(rbrace_rel) = body[rname_end..].find('{') else { break };
                let (r0, r1) = block_at(&body, rname_end + rbrace_rel);
                let rbody = body[r0..r1].to_string();

                for (fname, is_static) in free_functions(&rbody) {
                    let prefix = if is_static { "[static]" } else { "[method]" };
                    names.insert(format!("{prefix}{rname}.{fname}"));
                }
                if rbody
                    .lines()
                    .any(|l| l.trim_start().starts_with("constructor"))
                {
                    names.insert(format!("[constructor]{rname}"));
                }

                let blanked: String = std::iter::repeat(' ').take(r1 - scan.max(0)).collect();
                let _ = blanked;
                body.replace_range(rstart - "resource ".len()..=r1, &" ".repeat(r1 + 1 - (rstart - "resource ".len())));
                scan = r1;
            }

            for (fname, _) in free_functions(&body) {
                names.insert(fname);
            }
        }
    }
    // A guest EXPORT, never a host import — `command.wit` says `export run`.
    if let Some(run) = out.get_mut("run") {
        run.remove("run");
    }
    out.retain(|_, names| !names.is_empty());
    out
}

/// The parser must actually find something, in every package.
///
/// Without this a silently-failing parse would make every check below vacuous
/// — the exact failure mode this module exists to remove.
#[test]
fn the_wit_parser_finds_a_surface_in_every_package() {
    for package in PACKAGES {
        let surface = declared_surface(package);
        let count: usize = surface.values().map(|n| n.len()).sum();
        assert!(
            count > 0,
            "parsed no functions at all out of wasi:{package}'s WIT — the parser is broken, \
             and every assertion derived from it would pass vacuously"
        );
    }
}

/// Spot-check the parser against names read by eye from the WIT.
///
/// A parser checked only against its own output proves nothing. These four are
/// one of each shape it has to handle.
#[test]
fn the_wit_parser_recognises_all_four_shapes() {
    let sockets = declared_surface("sockets");
    let types = sockets.get("types").expect("wasi:sockets/types");
    assert!(
        types.contains("[static]tcp-socket.create"),
        "static not recognised"
    );
    assert!(
        types.contains("[method]tcp-socket.bind"),
        "method not recognised"
    );

    let http = declared_surface("http");
    let htypes = http.get("types").expect("wasi:http/types");
    assert!(
        htypes.contains("[constructor]fields"),
        "constructor not recognised"
    );

    let clocks = declared_surface("clocks");
    let mono = clocks
        .get("monotonic-clock")
        .expect("wasi:clocks/monotonic-clock");
    assert!(mono.contains("now"), "free function not recognised");
}

/// EVERY function the WIT declares is registered.
///
/// The same claim `interface_coverage::*_is_fully_registered` makes, but with
/// the declared side taken from the spec instead of a list. If the two ever
/// disagree, this one is right.
#[test]
fn every_declared_function_is_registered() {
    let registered = crate::interface_coverage::registered_wasi_names_for_test();
    let mut missing: Vec<String> = Vec::new();
    for package in PACKAGES {
        for (iface, names) in declared_surface(package) {
            let module = format!("wasi:{package}/{iface}");
            for name in names {
                if !registered.contains(&(module.clone(), name.clone())) {
                    missing.push(format!("{module} {name}"));
                }
            }
        }
    }
    assert!(
        missing.is_empty(),
        "{} functions are declared by the vendored WIT and NOT registered:\n  {}",
        missing.len(),
        missing.join("\n  ")
    );
}
