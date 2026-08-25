//! What the engine seam actually costs, measured rather than assumed.
//!
//! The question this answers: does routing every `web:*` call through a trait
//! object + an op enum make the DOM slow? It compares three paths over the
//! same work — a property write and read on a live control:
//!
//!   1. straight into `widgets` (the floor — no seam at all)
//!   2. through the seam (`DomOp` → `dyn WebEngine` → the same call)
//!   3. the same, but a node nested deep enough to exercise name lookup
//!
//! Run it with `--nocapture` to see the numbers. It asserts only on the ratio
//! between (1) and (2), which is the seam's own overhead — absolute timings
//! vary far too much between machines and build profiles to assert on.

use std::time::Instant;

use vybe_platform_web::engine::{DOCUMENT, DomOp, DomValue, apply, new_document};

const ITERS: u32 = 20_000;

fn create(doc: u64, tag: &str, input_type: &str) -> u64 {
    match apply(
        doc,
        DomOp::CreateElement {
            tag: tag.into(),
            input_type: input_type.into(),
        },
    ) {
        DomValue::Node(n) => n,
        other => panic!("expected a node, got {:?}", other),
    }
}

/// Seconds per operation, over `ITERS` write+read pairs.
fn time(label: &str, mut op: impl FnMut()) -> f64 {
    // Warm up so neither path pays first-touch costs.
    for _ in 0..1000 {
        op();
    }
    let start = Instant::now();
    for _ in 0..ITERS {
        op();
    }
    let per_op = start.elapsed().as_secs_f64() / ITERS as f64;
    println!("{label:<34} {:>8.0} ns/op", per_op * 1e9);
    per_op
}

#[test]
fn the_seam_is_not_where_the_cost_is() {
    vybe_platform_web::engine_widgets::install();
    let doc = new_document("bench");

    let flat = create(doc, "input", "text");
    apply(
        doc,
        DomOp::AppendChild {
            parent: DOCUMENT,
            child: flat,
        },
    );

    // A node five containers deep, so the by-name tree walk has real work.
    let mut parent = DOCUMENT;
    for _ in 0..5 {
        let div = create(doc, "div", "");
        apply(doc, DomOp::AppendChild { parent, child: div });
        parent = div;
    }
    let deep = create(doc, "input", "text");
    apply(
        doc,
        DomOp::AppendChild {
            parent,
            child: deep,
        },
    );

    println!();
    let direct = time("direct into widgets", || {
        vybe_platform_web::engine_widgets::with_document(doc, |d| {
            d.set_value(flat, "x");
            d.value(flat)
        });
    });

    let seam = time("through the seam", || {
        apply(doc, DomOp::SetValue(flat, "x".into()));
        apply(doc, DomOp::Value(flat));
    });

    let nested = time("through the seam, 6 levels deep", || {
        apply(doc, DomOp::SetValue(deep, "x".into()));
        apply(doc, DomOp::Value(deep));
    });
    println!();

    // The seam is a vtable hop plus an enum; the work underneath is a lock,
    // a name walk and a String. If dispatch ever dominates, this is where it
    // shows — and a generous bound still catches a real regression.
    let overhead = seam / direct;
    assert!(
        overhead < 3.0,
        "seam overhead {overhead:.2}x over the direct call — dispatch should \
         not dominate the work it dispatches ({:.0}ns vs {:.0}ns)",
        seam * 1e9,
        direct * 1e9
    );

    // Depth is the interesting axis: lookup is by NAME, so a deeper node
    // costs more. Recorded rather than bounded tightly — it is the number
    // that would justify switching to id-keyed dispatch.
    println!(
        "depth cost: {:.2}x going 6 levels deep\n",
        nested / seam.max(f64::MIN_POSITIVE)
    );
}
