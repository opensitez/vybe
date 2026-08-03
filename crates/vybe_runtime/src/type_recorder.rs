//! Optional VM instrumentation: records which `Value` variants flow
//! through each local slot during execution. The output drives the
//! anyref / ABI migration — if a slot only ever holds `I32`, a typed
//! WASM local can replace the boxed-externref representation and skip
//! the `wasm:js-number.{toI32,fromI32}` round-trip.
//!
//! The recorder is **off by default**. Enable via `VM::record_types(true)`
//! before `vm.run(...)`. Each `LOCAL_SET` / `LOCAL_GET` ticks a counter
//! for `(chunk_index, slot_index, value_tag)`; zero dispatch cost when
//! the recorder is `None`.
//!
//! After a run, `VM::take_type_record()` returns a printable summary
//! showing per-slot histograms and the fraction of slots that were
//! observed monomorphic — the upper bound on how much Phase 3 of the
//! migration can gain.

use crate::value::{Value, ValueTag};

/// Per-slot observation record. Index `i` is the count of times a value
/// tagged `ValueTag::<i>` was written to (or read from) this slot.
#[derive(Debug, Clone, Default)]
pub struct SlotObservations {
    pub counts: [u64; ValueTag::COUNT] }

impl SlotObservations {
    /// Number of distinct variants observed (0 = slot never written,
    /// 1 = monomorphic, ≥2 = polymorphic).
    pub fn distinct_variants(&self) -> usize {
        self.counts.iter().filter(|&&c| c > 0).count()
    }
    pub fn total(&self) -> u64 {
        self.counts.iter().sum()
    }
    /// If the slot is monomorphic, return the sole tag observed.
    pub fn monomorphic_tag(&self) -> Option<ValueTag> {
        let mut found: Option<ValueTag> = None;
        for (i, &c) in self.counts.iter().enumerate() {
            if c == 0 {
                continue;
            }
            if found.is_some() {
                return None;
            }
            found = Some(index_to_tag(i));
        }
        found
    }
}

fn index_to_tag(i: usize) -> ValueTag {
    match i {
        0 => ValueTag::Null,
        1 => ValueTag::Undefined,
        2 => ValueTag::Bool,
        3 => ValueTag::I32,
        4 => ValueTag::I64,
        5 => ValueTag::F64,
        6 => ValueTag::String,
        7 => ValueTag::Object,
        8 => ValueTag::WeakRef,
        9 => ValueTag::V128,
        10 => ValueTag::Symbol,
        11 => ValueTag::BigInt,
        _ => unreachable!() }
}

/// Per-VM observation bank. Sparse 2-D map: `slots[chunk_index][slot]`
/// grows lazily as execution encounters new slots.
#[derive(Debug, Default)]
pub struct TypeRecorder {
    slots: Vec<Vec<SlotObservations>> }

impl TypeRecorder {
    pub fn new() -> Self {
        Self { slots: Vec::new() }
    }

    /// Increment the counter for a value observed at `(chunk_index, slot)`.
    #[inline]
    pub fn record(&mut self, chunk_index: usize, slot: usize, value: &Value) {
        while self.slots.len() <= chunk_index {
            self.slots.push(Vec::new());
        }
        let row = &mut self.slots[chunk_index];
        while row.len() <= slot {
            row.push(SlotObservations::default());
        }
        row[slot].counts[value.tag().as_usize()] += 1;
    }

    /// Immutable view of the full observation bank.
    pub fn slots(&self) -> &[Vec<SlotObservations>] {
        &self.slots
    }

    /// Aggregate stats across every observed slot.
    pub fn summary(&self) -> TypeRecordSummary {
        let mut total = 0usize;
        let mut monomorphic = 0usize;
        let mut polymorphic = 0usize;
        let mut by_tag = [0u64; ValueTag::COUNT];
        for chunk in &self.slots {
            for slot in chunk {
                if slot.total() == 0 {
                    continue;
                }
                total += 1;
                if slot.distinct_variants() == 1 {
                    monomorphic += 1;
                } else {
                    polymorphic += 1;
                }
                for (i, &c) in slot.counts.iter().enumerate() {
                    by_tag[i] = by_tag[i].saturating_add(c);
                }
            }
        }
        TypeRecordSummary {
            observed_slots: total,
            monomorphic,
            polymorphic,
            by_tag }
    }

    /// Human-readable summary — useful to paste into an issue or
    /// migration notes. Shows per-chunk per-slot histograms plus an
    /// aggregate monomorphic-% line at the bottom.
    pub fn format_report(&self, chunk_names: &[String]) -> String {
        use std::fmt::Write;
        let mut out = String::new();
        for (ci, chunk_slots) in self.slots.iter().enumerate() {
            if chunk_slots.iter().all(|s| s.total() == 0) {
                continue;
            }
            let name = chunk_names.get(ci).map(String::as_str).unwrap_or("?");
            let _ = writeln!(&mut out, "\nchunk #{ci} ({name})");
            for (si, slot) in chunk_slots.iter().enumerate() {
                if slot.total() == 0 {
                    continue;
                }
                let _ = write!(&mut out, "  slot {si:3}: ");
                let mut first = true;
                for (i, &c) in slot.counts.iter().enumerate() {
                    if c == 0 {
                        continue;
                    }
                    if !first {
                        let _ = write!(&mut out, ", ");
                    }
                    first = false;
                    let _ = write!(&mut out, "{}={c}", index_to_tag(i).name());
                }
                if let Some(tag) = slot.monomorphic_tag() {
                    let _ = write!(&mut out, "  → mono({})", tag.name());
                } else {
                    let _ = write!(&mut out, "  → poly({})", slot.distinct_variants());
                }
                out.push('\n');
            }
        }
        let s = self.summary();
        let _ = writeln!(
            &mut out,
            "\nsummary: {} observed slots, {} mono ({:.1}%), {} poly",
            s.observed_slots,
            s.monomorphic,
            s.mono_percent(),
            s.polymorphic,
        );
        out
    }
}

/// Aggregate stats returned by `TypeRecorder::summary`.
#[derive(Debug, Clone)]
pub struct TypeRecordSummary {
    pub observed_slots: usize,
    pub monomorphic: usize,
    pub polymorphic: usize,
    pub by_tag: [u64; ValueTag::COUNT] }

impl TypeRecordSummary {
    pub fn mono_percent(&self) -> f64 {
        if self.observed_slots == 0 {
            0.0
        } else {
            100.0 * self.monomorphic as f64 / self.observed_slots as f64
        }
    }
}
