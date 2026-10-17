use vybe_ast::word_local_transport::WordLocalChunkProof;
use vybe_ast::{Directives, WordLocalTransport};

#[test]
fn raw_word_local_transport_is_never_the_unspecified_default() {
    let policy = Directives::default();
    assert_eq!(policy.word_local_transport, None);
    assert_eq!(policy.word_local_transport(), WordLocalTransport::Boxed);
}

#[test]
fn unspecified_child_inherits_and_explicit_boxed_child_disables_proof_analysis() {
    let mut policy = Directives {
        word_local_transport: Some(WordLocalTransport::ProvenW32),
        ..Default::default()
    };
    policy.overlay(&Directives::default());
    assert_eq!(policy.word_local_transport(), WordLocalTransport::ProvenW32);
    policy.overlay(&Directives {
        word_local_transport: Some(WordLocalTransport::Boxed),
        ..Default::default()
    });
    assert_eq!(policy.word_local_transport(), WordLocalTransport::Boxed);
}

#[test]
fn backend_sidecar_identity_rejects_rebased_chunks_and_changed_code() {
    let proof = WordLocalChunkProof {
        chunk_index: 2,
        code: vec![1, 2, 3],
        locals: vec![],
    };
    assert!(proof.matches_code(2, &[1, 2, 3]));
    assert!(!proof.matches_code(3, &[1, 2, 3]));
    assert!(!proof.matches_code(2, &[1, 2, 4]));
    assert!(!proof.matches_code(2, &[1, 2]));
}
