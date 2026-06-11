use std::sync::Arc;
use vybe_bytecode::{Chunk, Op, VM, Value};
use vybe_host::{Capabilities, register_with_capabilities};

fn call_import(module: &str, name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<wasi-crypto-test>");
    let import_idx = chunk.add_import(module, name);
    let argc = args.len() as u8;
    for value in args {
        let constant = chunk.add_constant(value);
        chunk.emit_op_u16(Op::CONST, constant, 0);
    }
    chunk.emit_op_u16(Op::CALL_IMPORT, import_idx, 0);
    chunk.emit(argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn has_import(module: &str, name: &str) -> bool {
    let mut vm = VM::new();
    register_with_capabilities(&mut vm, &Capabilities::all());
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

#[test]
fn proposal_wasi_crypto_common_surface_is_registered() {
    let expected = [
        "options_open",
        "options_close",
        "options_set",
        "options_set_u64",
        "options_set_guest_buffer",
        "array_output_len",
        "array_output_pull",
        "secrets_manager_open",
        "secrets_manager_close",
        "secrets_manager_invalidate",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import("wasi_ephemeral_crypto_common", name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing wasi-crypto common imports: {missing:?}"
    );
}

#[test]
fn proposal_wasi_crypto_asymmetric_common_surface_is_registered() {
    let expected = [
        "keypair_generate",
        "keypair_import",
        "keypair_generate_managed",
        "keypair_store_managed",
        "keypair_replace_managed",
        "keypair_id",
        "keypair_from_id",
        "keypair_from_pk_and_sk",
        "keypair_export",
        "keypair_publickey",
        "keypair_secretkey",
        "keypair_close",
        "publickey_import",
        "publickey_export",
        "publickey_verify",
        "publickey_from_secretkey",
        "publickey_close",
        "secretkey_import",
        "secretkey_export",
        "secretkey_close",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import("wasi_ephemeral_crypto_asymmetric_common", name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing wasi-crypto asymmetric-common imports: {missing:?}"
    );
}

#[test]
fn proposal_wasi_crypto_symmetric_surface_is_registered() {
    let expected = [
        "symmetric_key_generate",
        "symmetric_key_import",
        "symmetric_key_export",
        "symmetric_key_close",
        "symmetric_key_generate_managed",
        "symmetric_key_store_managed",
        "symmetric_key_replace_managed",
        "symmetric_key_id",
        "symmetric_key_from_id",
        "symmetric_state_open",
        "symmetric_state_options_get",
        "symmetric_state_options_get_u64",
        "symmetric_state_clone",
        "symmetric_state_close",
        "symmetric_state_absorb",
        "symmetric_state_squeeze",
        "symmetric_state_squeeze_tag",
        "symmetric_state_squeeze_key",
        "symmetric_state_max_tag_len",
        "symmetric_state_encrypt",
        "symmetric_state_encrypt_detached",
        "symmetric_state_decrypt",
        "symmetric_state_decrypt_detached",
        "symmetric_state_ratchet",
        "symmetric_tag_len",
        "symmetric_tag_pull",
        "symmetric_tag_verify",
        "symmetric_tag_close",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import("wasi_ephemeral_crypto_symmetric", name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing wasi-crypto symmetric imports: {missing:?}"
    );
}

#[test]
fn proposal_wasi_crypto_signatures_surface_is_registered() {
    let expected = [
        "signature_export",
        "signature_import",
        "signature_state_open",
        "signature_state_update",
        "signature_state_sign",
        "signature_state_close",
        "signature_verification_state_open",
        "signature_verification_state_update",
        "signature_verification_state_verify",
        "signature_verification_state_close",
        "signature_close",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import("wasi_ephemeral_crypto_signatures", name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing wasi-crypto signatures imports: {missing:?}"
    );
}

#[test]
fn proposal_wasi_crypto_signatures_batch_surface_is_registered() {
    let expected = ["batch_signature_state_sign", "batch_signature_state_verify"];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import("wasi_ephemeral_crypto_signatures_batch", name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing wasi-crypto signatures-batch imports: {missing:?}"
    );
}

#[test]
fn proposal_wasi_crypto_kx_surface_is_registered() {
    let expected = ["kx_dh", "kx_encapsulate", "kx_decapsulate"];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import("wasi_ephemeral_crypto_kx", name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing wasi-crypto key-exchange imports: {missing:?}"
    );
}

#[test]
fn proposal_wasi_crypto_external_secrets_surface_is_registered() {
    let expected = [
        "external_secret_store",
        "external_secret_replace",
        "external_secret_from_id",
        "external_secret_invalidate",
        "external_secret_encapsulate",
        "external_secret_decapsulate",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import("wasi_ephemeral_crypto_external_secrets", name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing wasi-crypto external-secrets imports: {missing:?}"
    );
}

#[test]
fn proposal_wasi_crypto_symmetric_batch_surface_is_registered() {
    let expected = [
        "batch_symmetric_state_squeeze",
        "batch_symmetric_state_squeeze_tag",
        "batch_symmetric_state_encrypt",
        "batch_symmetric_state_encrypt_detached",
        "batch_symmetric_state_decrypt",
        "batch_symmetric_state_decrypt_detached",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import("wasi_ephemeral_crypto_symmetric_batch", name))
        .collect::<Vec<_>>();
    assert!(
        missing.is_empty(),
        "missing wasi-crypto symmetric-batch imports: {missing:?}"
    );
}
