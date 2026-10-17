use std::sync::Arc;
use vybe_runtime::{Chunk, Op, VM};

fn module(names: &[&str], bodies: usize) -> Vec<Chunk> {
    let globals = Arc::new(names.iter().map(|name| name.to_string())
        .collect::<vybe_runtime::segmented_global_names::SegmentedGlobalNames>());
    (0..bodies).map(|index| {
        let mut chunk = Chunk::new(&format!("global-sharing-{index}"));
        chunk.globals = Arc::clone(&globals);
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        chunk.emit_op(Op::RETURN, 0);
        chunk
    }).collect()
}

#[test]
fn sharing_preserves_incompatible_tables_and_interleaved_snapshots() {
    let mut vm = VM::new();
    vm.run_linked(module(&["first", "second"], 8), vec![]).unwrap();
    let original = Arc::clone(&vm.chunks[0].globals);
    let alias = Arc::new(vec!["different-name".to_string()].into());
    let hole = Arc::new(vec![String::new()].into());
    // Repeated and interleaved identities exercise both the last-table fast
    // path and the distinct-table fallback, including negative prefix proofs.
    vm.chunks[1].globals = Arc::clone(&alias);
    vm.chunks[2].globals = Arc::clone(&alias);
    vm.chunks[4].globals = Arc::clone(&alias);
    vm.chunks[6].globals = Arc::clone(&hole);

    vm.run_linked(module(&["third"], 3), vec![]).unwrap();
    let authoritative = Arc::clone(&vm.chunks[8].globals);
    assert_eq!(&authoritative[..], &["undefined", "first", "second", "third"]);
    for index in [0, 8, 9, 10] {
        assert!(Arc::ptr_eq(&vm.chunks[index].globals, &authoritative));
    }
    for index in [3, 5, 7] {
        assert!(Arc::ptr_eq(&vm.chunks[index].globals, &original));
        assert!(vm.chunks[index].globals.is_prefix_of(&authoritative));
    }
    for index in [1, 2, 4] {
        assert!(Arc::ptr_eq(&vm.chunks[index].globals, &alias));
    }
    assert!(Arc::ptr_eq(&vm.chunks[6].globals, &hole));
    assert_eq!(&original[..], &["undefined", "first", "second"]);
}

#[test]
#[ignore = "focused native linkage benchmark; run explicitly"]
fn repeated_module_linkage_benchmark() {
    let mut vm = VM::new();
    let started = std::time::Instant::now();
    for index in 0..300 {
        let name = format!("global-{index}");
        vm.run_linked(module(&[&name], 24), vec![]).unwrap();
    }
    assert_eq!(vm.chunks.len(), 7_200);
    assert_eq!(vm.chunks.last().unwrap().globals.len(), 301);
    eprintln!("linked_modules=300 chunks=7200 elapsed_ms={:.6}",
        started.elapsed().as_secs_f64() * 1000.0);
}
