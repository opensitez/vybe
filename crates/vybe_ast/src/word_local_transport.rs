//! Optional backend proof data; it never changes the source value ABI itself.

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum WordLocalTransport {
    #[default]
    Boxed,
    ProvenW32,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WordInterpretation {
    Signed,
    Unsigned,
}

/// The value representation at an independently validated producer/consumer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WordBoundaryRepresentation {
    BoxedSignedW32,
    BoxedUnsignedW32,
    NativeW32,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct WordBoundaryProof {
    /// Instruction offset in the exact finalized bytecode snapshot.
    pub offset: usize,
    pub representation: WordBoundaryRepresentation,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum WordLocalAccessProof {
    Read {
        offset: usize,
        consumer: WordBoundaryProof,
    },
    Write {
        offset: usize,
        producer: WordBoundaryProof,
    },
    Tee {
        offset: usize,
        producer: WordBoundaryProof,
        consumer: WordBoundaryProof,
    },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WordLocalSlotProof {
    /// Final bytecode slot, not a Wasm index. The writer applies its ABI map.
    pub slot: u16,
    pub interpretation: WordInterpretation,
    /// Complete slot-wide accesses, including compiler scratch reuse.
    pub accesses: Vec<WordLocalAccessProof>,
}

/// An advisory sidecar. Consumers must validate code identity, complete slot
/// accesses, producer types, initialization and CFG joins before selecting raw
/// storage. A match alone is not authorization; stale/invalid proofs stay boxed.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WordLocalChunkProof {
    pub chunk_index: usize,
    /// Exact identity, rather than a collision-prone hash or a source timestamp.
    pub code: Vec<u8>,
    pub locals: Vec<WordLocalSlotProof>,
}

impl WordLocalChunkProof {
    pub fn matches_code(&self, chunk_index: usize, code: &[u8]) -> bool {
        self.chunk_index == chunk_index && self.code == code
    }
}
