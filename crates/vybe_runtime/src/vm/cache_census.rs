//! Explicitly requested, read-only execution-cache allocation accounting.
//!
//! These are Rust Vec capacities, not allocator usable sizes or RSS. Nested
//! allocations owned by cache entries and shared block-target caches are not
//! included. Nothing invokes this census on an execution hot path.

use std::fmt::Write;
use std::mem::size_of;

use super::{LocalFastPath, NativeResultFastPath, VM};

fn write_rows<T>(
    output: &mut String,
    name: &str,
    rows: &Vec<Vec<T>>,
    populated: Option<fn(&T) -> bool>,
) -> usize {
    let element_bytes = size_of::<T>();
    let live_slots = rows.iter().map(Vec::len).sum::<usize>();
    let capacity_slots = rows.iter().map(Vec::capacity).sum::<usize>();
    let live_bytes = live_slots.saturating_mul(element_bytes);
    let capacity_bytes = capacity_slots.saturating_mul(element_bytes);
    let headers = rows.capacity().saturating_mul(size_of::<Vec<T>>());
    let nonempty = rows.iter().filter(|row| !row.is_empty()).count();
    let _ = write!(output,
        "cache_table name={name} element_bytes={element_bytes} rows={} nonempty_rows={nonempty} live_slots={live_slots} capacity_slots={capacity_slots} live_bytes={live_bytes} capacity_bytes={capacity_bytes} outer_capacity_bytes={headers}",
        rows.len());
    if let Some(is_populated) = populated {
        let count = rows.iter().flatten().filter(|entry| is_populated(entry)).count();
        let _ = write!(output, " populated_slots={count}");
    }
    output.push('\n');
    capacity_bytes.saturating_add(headers)
}

fn row_capacity_bytes<T>(rows: &[Vec<T>], index: usize) -> usize {
    rows.get(index)
        .map_or(0, |row| row.capacity().saturating_mul(size_of::<T>()))
}

impl VM {
    /// Account for the retained capacity of execution metadata on explicit
    /// diagnostic request. Does not populate, resize, clear, or invalidate any
    /// cache. Occupancy counting scans the native/local cache rows, so callers
    /// should request this at a paused or completed execution boundary.
    pub(super) fn execution_cache_census_base(&self) -> String {
        let mut output = String::new();
        let code_bytes = self.chunks.iter().map(|chunk| chunk.code.len()).sum::<usize>();
        let code_capacity_bytes = self.chunks.iter()
            .map(|chunk| chunk.code.capacity()).sum::<usize>();
        let _ = writeln!(output,
            "cache_census chunks={} code_bytes={code_bytes} code_capacity_bytes={code_capacity_bytes} excludes_nested_entry_allocations=true excludes_shared_block_targets=true",
            self.chunks.len());
        let mut total = 0usize;
        macro_rules! table {
            ($field:ident, $populated:expr) => {
                total = total.saturating_add(write_rows(
                    &mut output, stringify!($field), &self.$field, $populated));
            };
        }
        table!(decoded_ops, None);
        table!(decoded_operands, None);
        table!(decoded_operands64, None);
        table!(chunk_import_cache, Some(|entry| entry.is_some()));
        table!(chunk_call_target_cache, Some(|entry| entry.is_some()));
        table!(native_call_sites, Some(|entry| entry.is_some()));
        table!(native_result_fast_paths,
            Some(|entry| !matches!(entry, NativeResultFastPath::None)));
        table!(local_fast_paths, Some(|entry| !matches!(entry, LocalFastPath::None)));
        table!(loop_superops, None);
        table!(br_table_targets, Some(|entry| entry.is_some()));
        let pending_sites = self.pending_native_call_sites.values()
            .map(Vec::len).sum::<usize>();
        let pending_capacity_bytes = self.pending_native_call_sites.values()
            .map(|row| row.capacity().saturating_mul(size_of::<(usize, super::NativeCallSiteTarget)>()))
            .sum::<usize>();
        let _ = writeln!(output,
            "cache_pending_native_calls rows={} sites={pending_sites} vector_capacity_bytes={pending_capacity_bytes} hash_map_capacity={} excludes_hash_map_storage=true",
            self.pending_native_call_sites.len(), self.pending_native_call_sites.capacity());
        total = total.saturating_add(pending_capacity_bytes);
        let _ = writeln!(output, "cache_census capacity_bytes_including_outer_tables={total}");

        let mut largest: Vec<_> = (0..self.chunks.len()).map(|index| {
            let sites = row_capacity_bytes(&self.native_call_sites, index);
            let results = row_capacity_bytes(&self.native_result_fast_paths, index);
            let locals = row_capacity_bytes(&self.local_fast_paths, index);
            (sites.saturating_add(results).saturating_add(locals), index, sites, results, locals)
        }).filter(|row| row.0 != 0).collect();
        largest.sort_unstable_by(|left, right|
            right.0.cmp(&left.0).then_with(|| left.1.cmp(&right.1)));
        for (bytes, index, sites, results, locals) in largest.into_iter().take(20) {
            let chunk = &self.chunks[index];
            let _ = writeln!(output,
                "cache_chunk index={index} name={:?} code_bytes={} native_local_capacity_bytes={bytes} native_call_capacity_bytes={sites} native_result_capacity_bytes={results} local_capacity_bytes={locals}",
                chunk.name, chunk.code.len());
        }
        output
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn row_census_counts_capacity_and_occupancy_without_modification() {
        let mut populated = Vec::with_capacity(64);
        populated.extend([None, Some(7u32), None]);
        let rows = vec![populated, Vec::new()];
        let pointer = rows[0].as_ptr();
        let capacity = rows[0].capacity();
        let outer_capacity = rows.capacity();
        let mut report = String::new();
        let bytes = write_rows(&mut report, "probe", &rows, Some(Option::is_some));
        assert_eq!(bytes, capacity * size_of::<Option<u32>>()
            + outer_capacity * size_of::<Vec<Option<u32>>>());
        assert!(report.contains("rows=2 nonempty_rows=1 live_slots=3"));
        assert!(report.contains("populated_slots=1"));
        assert_eq!(rows[0].as_ptr(), pointer);
        assert_eq!(rows[0].capacity(), capacity);
        assert_eq!(rows[0], [None, Some(7), None]);
    }
}
