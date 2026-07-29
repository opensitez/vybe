//! Coverage audit for every proposal directory vendored under
//! `proposals/spec/proposals`.
//!
//! This is intentionally a manifest test: it does not prove semantic
//! completeness by itself, but it prevents silent drift between the upstream
//! proposal set and Vybe's local compliance tests.

use std::collections::BTreeSet;
use std::path::{Path, PathBuf};

struct ProposalCoverage {
    proposal: &'static str,
    tests: &'static [&'static str],
    upstream_wast_dirs: &'static [&'static str],
}

const PROPOSAL_COVERAGE: &[ProposalCoverage] = &[
    ProposalCoverage {
        proposal: "annotations",
        tests: &["wasm_compliance_test.rs"],
        upstream_wast_dirs: &["test/core/annotations.wast"],
    },
    ProposalCoverage {
        proposal: "branch-hinting",
        tests: &["branch_hinting_test.rs", "compilation_hints_test.rs"],
        upstream_wast_dirs: &["test/custom/metadata.code.branch_hint"],
    },
    ProposalCoverage {
        proposal: "bulk-memory-operations",
        tests: &["bulk_memory_operations_test.rs"],
        upstream_wast_dirs: &["test/core/bulk-memory"],
    },
    ProposalCoverage {
        proposal: "exception-handling",
        tests: &["exception_handling_test.rs", "wasm_compliance_test.rs"],
        upstream_wast_dirs: &[
            "test/core/exceptions",
            "test/js-api/exception",
            "test/js-api/tag",
        ],
    },
    ProposalCoverage {
        proposal: "extended-const",
        tests: &["extended_const_test.rs"],
        upstream_wast_dirs: &[],
    },
    ProposalCoverage {
        proposal: "function-references",
        tests: &["function_references_test.rs", "wasm_test.rs"],
        upstream_wast_dirs: &[
            "test/core/call_ref.wast",
            "test/core/ref_func.wast",
            "test/core/return_call_ref.wast",
        ],
    },
    ProposalCoverage {
        proposal: "gc",
        tests: &[
            "gc_test.rs",
            "vm_type_system_test.rs",
            "wasm_compliance_test.rs",
        ],
        upstream_wast_dirs: &["test/core/gc", "test/js-api/gc"],
    },
    ProposalCoverage {
        proposal: "js-string-builtins",
        tests: &["js_builtins_compliance_test.rs", "wasm_test.rs"],
        upstream_wast_dirs: &["test/js-api/js-string"],
    },
    ProposalCoverage {
        proposal: "memory64",
        tests: &["wasm_binary_format_test.rs", "wasm_test.rs"],
        upstream_wast_dirs: &["test/core/memory64"],
    },
    ProposalCoverage {
        proposal: "multi-memory",
        tests: &["multi_memory_test.rs", "wasm_binary_format_test.rs"],
        upstream_wast_dirs: &["test/core/multi-memory"],
    },
    ProposalCoverage {
        proposal: "multi-value",
        tests: &["wasm_compliance_test.rs", "wasm_structured_control_test.rs"],
        upstream_wast_dirs: &[],
    },
    ProposalCoverage {
        proposal: "nontrapping-float-to-int-conversion",
        tests: &[
            "nontrapping_float_to_int_test.rs",
            "wasm_compliance_test.rs",
        ],
        upstream_wast_dirs: &[],
    },
    ProposalCoverage {
        proposal: "reference-types",
        tests: &["reference_types_test.rs", "wasm_compliance_test.rs"],
        upstream_wast_dirs: &[
            "test/core/ref.wast",
            "test/core/ref_as_non_null.wast",
            "test/core/ref_is_null.wast",
            "test/core/ref_null.wast",
        ],
    },
    ProposalCoverage {
        proposal: "relaxed-simd",
        tests: &["relaxed_simd_test.rs", "wasm_compliance_test.rs"],
        upstream_wast_dirs: &["test/core/relaxed-simd"],
    },
    ProposalCoverage {
        proposal: "sign-extension-ops",
        tests: &["sign_extension_ops_test.rs", "wasm_test.rs"],
        upstream_wast_dirs: &[],
    },
    ProposalCoverage {
        proposal: "simd",
        tests: &["simd_test.rs", "wasm_test.rs"],
        upstream_wast_dirs: &["test/core/simd"],
    },
    ProposalCoverage {
        proposal: "tail-call",
        tests: &["tail_call_test.rs", "wasm_compliance_test.rs"],
        upstream_wast_dirs: &[
            "test/core/return_call.wast",
            "test/core/return_call_indirect.wast",
            "test/core/return_call_ref.wast",
        ],
    },
];

fn repo_root() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .and_then(Path::parent)
        .expect("vybe_runtime should be under crates/")
        .to_path_buf()
}

#[test]
fn every_upstream_proposal_has_explicit_local_test_mapping() {
    let proposals_dir = repo_root().join("proposals/spec/proposals");
    let found: BTreeSet<String> = std::fs::read_dir(&proposals_dir)
        .unwrap_or_else(|err| panic!("failed to read {}: {err}", proposals_dir.display()))
        .filter_map(Result::ok)
        .filter(|entry| entry.file_type().map(|ty| ty.is_dir()).unwrap_or(false))
        .map(|entry| entry.file_name().to_string_lossy().into_owned())
        .collect();

    let mapped: BTreeSet<String> = PROPOSAL_COVERAGE
        .iter()
        .map(|coverage| coverage.proposal.to_string())
        .collect();

    assert_eq!(
        found, mapped,
        "every proposals/spec/proposals entry must be explicitly mapped to local tests"
    );
}

#[test]
fn mapped_proposal_test_files_exist() {
    // Proposal suites live in one of two places: this crate's own `tests/`
    // (VM-level semantics) or `platforms/wasm/tests/` (binary reader/writer),
    // since the `.wasm` codec was split out into its own platform crate.
    let manifest = Path::new(env!("CARGO_MANIFEST_DIR"));
    let search_dirs = [
        manifest.join("tests"),
        manifest
            .join("..")
            .join("..")
            .join("platforms")
            .join("wasm")
            .join("tests"),
    ];
    for coverage in PROPOSAL_COVERAGE {
        assert!(
            !coverage.tests.is_empty(),
            "{} must name at least one local test file",
            coverage.proposal
        );
        for test in coverage.tests {
            assert!(
                search_dirs.iter().any(|dir| dir.join(test).is_file()),
                "{} maps to missing test file {test} (looked in {})",
                coverage.proposal,
                search_dirs
                    .iter()
                    .map(|d| d.display().to_string())
                    .collect::<Vec<_>>()
                    .join(", ")
            );
        }
    }
}

#[test]
fn mapped_upstream_wast_locations_exist_when_listed() {
    let spec_dir = repo_root().join("proposals/spec");
    for coverage in PROPOSAL_COVERAGE {
        for relative in coverage.upstream_wast_dirs {
            let path = spec_dir.join(relative);
            assert!(
                path.exists(),
                "{} maps to missing upstream spec-test location {}",
                coverage.proposal,
                path.display()
            );
        }
    }
}
