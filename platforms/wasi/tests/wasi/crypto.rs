use std::sync::Arc;
use vybe_compiler::primitives::platforms::register_platforms;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::{Chunk, Op, VM, Value};

fn call_import(module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<wasi-crypto-test>");
    let import_idx = chunk.add_import(module, name);
    let argc = args.len() as u8;
    for value in args {
        match value {
            Value::I32(n) => chunk.emit_i32_const(n, 0),
            Value::I64(n) => chunk.emit_i64_const(n, 0),
            Value::F32(f) => chunk.emit_f32_const(f, 0),
            Value::F64(f) => chunk.emit_f64_const(f, 0),
            Value::Bool(b) => chunk.emit_bool_const(b, 0),
            Value::String(text) => chunk.emit_string_const(&text, 0),
            Value::Null => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0),
            other => panic!("no spec const emitter for test argument {other:?}"),
        }
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn has_import(module: &str, name: &str) -> bool {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.host_registry
        .contains_key(&(module.to_string(), name.to_string()))
}

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

#[test]
fn sha256_of_empty_string_matches_known_digest() {
    let digest = call_import("wasi:crypto/hashes", "sha256", vec![s("")]);
    assert_eq!(
        digest,
        s("e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855")
    );
}

#[test]
fn sha256_of_ascii_payload_matches_known_digest() {
    let digest = call_import("wasi:crypto/hashes", "sha256", vec![s("abc")]);
    assert_eq!(
        digest,
        s("ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad")
    );
}

#[test]
fn md5_of_empty_string_matches_known_digest() {
    let digest = call_import("wasi:crypto/hashes", "md5", vec![s("")]);
    assert_eq!(digest, s("d41d8cd98f00b204e9800998ecf8427e"));
}

#[test]
fn md5_of_ascii_payload_matches_known_digest() {
    let digest = call_import("wasi:crypto/hashes", "md5", vec![s("abc")]);
    assert_eq!(digest, s("900150983cd24fb0d6963f7d28e17f72"));
}

#[test]
fn hash_algorithms_produce_distinct_digests_for_same_payload() {
    let sha = call_import("wasi:crypto/hashes", "sha256", vec![s("vybe")]);
    let md5 = call_import("wasi:crypto/hashes", "md5", vec![s("vybe")]);
    assert_ne!(sha, md5);
}

/// ⚠BOTH SIDES OF THIS WERE WRONG, AND EACH HID THE OTHER.
///
/// The spec is `proposals/wasi-crypto/wit/`: `package wasi:crypto@0.11.0`,
/// `interface wasi-ephemeral-crypto-common`, kebab-case function names. So the
/// module is `wasi:crypto/wasi-ephemeral-crypto-common`.
///
/// This file asserted the WITX module (`wasi_ephemeral_crypto_common`, snake
/// names) and the host registered an INVENTED short one (`wasi:crypto/common`).
/// Neither is a name a guest generated from this WIT would import. And because
/// the test was wrong too, it reported every function missing — so "the module
/// name is invented" was indistinguishable from "the interface is absent".
///
/// The ten names below are the WIT's whole `common` surface, in declaration
/// order, including `options-set-u64` — which the host really does not
/// register. That is now the only thing this test can be red about.
#[test]
fn proposal_wasi_crypto_common_surface_is_registered() {
    let expected = [
        "options-open",
        "options-close",
        "options-set",
        "options-set-u64",
        "options-set-guest-buffer",
        "array-output-len",
        "array-output-pull",
        "secrets-manager-open",
        "secrets-manager-close",
        "secrets-manager-invalidate",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import("wasi:crypto/wasi-ephemeral-crypto-common", name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing wasi-crypto common imports: {missing:?}"
    );
}

#[test]
fn proposal_wasi_crypto_asymmetric_common_surface_is_registered() {
    let expected = [
        "keypair-generate",
        "keypair-import",
        "keypair-generate-managed",
        "keypair-store-managed",
        "keypair-replace-managed",
        "keypair-id",
        "keypair-from-id",
        "keypair-from-pk-and-sk",
        "keypair-export",
        "keypair-publickey",
        "keypair-secretkey",
        "keypair-close",
        "publickey-import",
        "publickey-export",
        "publickey-verify",
        "publickey-from-secretkey",
        "publickey-close",
        "secretkey-import",
        "secretkey-export",
        "secretkey-close",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import("wasi:crypto/wasi-ephemeral-crypto-asymmetric-common", name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing wasi-crypto asymmetric-common imports: {missing:?}"
    );
}

#[test]
fn proposal_wasi_crypto_symmetric_surface_is_registered() {
    let expected = [
        "symmetric-key-generate",
        "symmetric-key-import",
        "symmetric-key-export",
        "symmetric-key-close",
        "symmetric-key-generate-managed",
        "symmetric-key-store-managed",
        "symmetric-key-replace-managed",
        "symmetric-key-id",
        "symmetric-key-from-id",
        "symmetric-state-open",
        "symmetric-state-options-get",
        "symmetric-state-options-get-u64",
        "symmetric-state-clone",
        "symmetric-state-close",
        "symmetric-state-absorb",
        "symmetric-state-squeeze",
        "symmetric-state-squeeze-tag",
        "symmetric-state-squeeze-key",
        "symmetric-state-max-tag-len",
        "symmetric-state-encrypt",
        "symmetric-state-encrypt-detached",
        "symmetric-state-decrypt",
        "symmetric-state-decrypt-detached",
        "symmetric-state-ratchet",
        "symmetric-tag-len",
        "symmetric-tag-pull",
        "symmetric-tag-verify",
        "symmetric-tag-close",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import("wasi:crypto/wasi-ephemeral-crypto-symmetric", name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing wasi-crypto symmetric imports: {missing:?}"
    );
}

#[test]
fn proposal_wasi_crypto_signatures_surface_is_registered() {
    let expected = [
        "signature-export",
        "signature-import",
        "signature-state-open",
        "signature-state-update",
        "signature-state-sign",
        "signature-state-close",
        "signature-verification-state-open",
        "signature-verification-state-update",
        "signature-verification-state-verify",
        "signature-verification-state-close",
        "signature-close",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import("wasi:crypto/wasi-ephemeral-crypto-signatures", name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing wasi-crypto signatures imports: {missing:?}"
    );
}

#[test]
fn proposal_wasi_crypto_signatures_batch_surface_is_registered() {
    let expected = ["batch-signature-state-sign", "batch-signature-state-verify"];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import("wasi:crypto/wasi-ephemeral-crypto-signatures-batch", name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing wasi-crypto signatures-batch imports: {missing:?}"
    );
}

#[test]
fn proposal_wasi_crypto_kx_surface_is_registered() {
    let expected = ["kx-dh", "kx-encapsulate", "kx-decapsulate"];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import("wasi:crypto/wasi-ephemeral-crypto-kx", name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing wasi-crypto key-exchange imports: {missing:?}"
    );
}

#[test]
fn proposal_wasi_crypto_external_secrets_surface_is_registered() {
    let expected = [
        "external-secret-store",
        "external-secret-replace",
        "external-secret-from-id",
        "external-secret-invalidate",
        "external-secret-encapsulate",
        "external-secret-decapsulate",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import("wasi:crypto/wasi-ephemeral-crypto-external-secrets", name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing wasi-crypto external-secrets imports: {missing:?}"
    );
}

#[test]
fn proposal_wasi_crypto_symmetric_batch_surface_is_registered() {
    let expected = [
        "batch-symmetric-state-squeeze",
        "batch-symmetric-state-squeeze-tag",
        "batch-symmetric-state-encrypt",
        "batch-symmetric-state-encrypt-detached",
        "batch-symmetric-state-decrypt",
        "batch-symmetric-state-decrypt-detached",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import("wasi:crypto/wasi-ephemeral-crypto-symmetric-batch", name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing wasi-crypto symmetric-batch imports: {missing:?}"
    );
}
