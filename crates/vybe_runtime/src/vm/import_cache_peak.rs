// Opt-in link-time import-cache capacity census. Recorded before execution
// can clear the current unit's caches. This is not an allocator/RSS peak.

use super::VM;
use std::fmt::Write as _;

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct CacheStorage {
    pub rows: usize,
    pub nonempty_rows: usize,
    pub capacity_slots: usize,
    pub capacity_bytes: usize,
    pub populated_slots: usize,
}

fn storage<T>(rows: &Vec<Vec<Option<Option<T>>>>) -> CacheStorage {
    let mut result = CacheStorage {
        rows: rows.len(),
        capacity_bytes: rows.capacity().saturating_mul(
            std::mem::size_of::<Vec<Option<Option<T>>>>(),
        ),
        ..CacheStorage::default()
    };
    for row in rows {
        result.nonempty_rows += usize::from(!row.is_empty());
        result.capacity_slots = result.capacity_slots.saturating_add(row.capacity());
        result.capacity_bytes = result.capacity_bytes.saturating_add(
            row.capacity().saturating_mul(std::mem::size_of::<Option<Option<T>>>()),
        );
    }
    result
}

fn populated<T>(rows: &[Vec<Option<Option<T>>>]) -> usize {
    // Some(None) is a populated cached miss, not an uninitialized slot.
    rows.iter().map(|row| row.iter().filter(|entry| entry.is_some()).count()).sum()
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct ImportCachePeak {
    pub first_chunk: usize,
    pub end_chunk: usize,
    pub imports: CacheStorage,
    pub call_targets: CacheStorage,
    pub combined_capacity_bytes: usize,
}

#[derive(Default)]
pub struct ImportCachePeakRecorder {
    pub enabled: bool,
    pub peak: Option<ImportCachePeak>,
}

impl ImportCachePeakRecorder {
    pub fn record<T, U>(
        &mut self,
        imports: &Vec<Vec<Option<Option<T>>>>,
        call_targets: &Vec<Vec<Option<Option<U>>>>,
        first_chunk: usize,
        end_chunk: usize,
    ) -> bool {
        if !self.enabled {
            return false;
        }
        let mut import_storage = storage(imports);
        let mut target_storage = storage(call_targets);
        let combined = import_storage.capacity_bytes.saturating_add(target_storage.capacity_bytes);
        if self.peak.is_some_and(|peak| peak.combined_capacity_bytes >= combined) {
            return false;
        }
        // Scan occupancy only for a new peak. No allocation, cloning of cache
        // entries, or formatting is needed on either path.
        import_storage.populated_slots = populated(imports);
        target_storage.populated_slots = populated(call_targets);
        self.peak = Some(ImportCachePeak {
            first_chunk,
            end_chunk,
            imports: import_storage,
            call_targets: target_storage,
            combined_capacity_bytes: combined,
        });
        true
    }
}

impl VM {
    #[inline]
    pub(super) fn record_import_cache_peak(&mut self, first_chunk: usize) {
        self.import_cache_peak.record(
            &self.chunk_import_cache,
            &self.chunk_call_target_cache,
            first_chunk,
            self.chunks.len(),
        );
    }

    /// Includes a historical pre-execution cache-capacity peak when the VM
    /// was constructed with VYBE_VM_CACHE_CENSUS set. Reset/reload retain this
    /// numeric receipt; its chunk range describes that historical link, not
    /// a live chunk handle. Child VMs construct their own empty recorder.
    pub fn execution_cache_census(&self) -> String {
        let mut out = self.execution_cache_census_base();
        if !self.import_cache_peak.enabled {
            return out;
        }
        if let Some(peak) = self.import_cache_peak.peak {
            let _ = writeln!(out,
                "prefill import-cache peak: combined_capacity_bytes={} unit_chunks={}..{} (historical; excludes entries' nested allocations)",
                peak.combined_capacity_bytes, peak.first_chunk, peak.end_chunk);
            for (name, row) in [
                ("chunk_import_cache", peak.imports),
                ("chunk_call_target_cache", peak.call_targets),
            ] {
                let _ = writeln!(out,
                    "  {}: rows={} nonempty_rows={} capacity_slots={} populated_slots={} capacity_bytes={}",
                    name, row.rows, row.nonempty_rows, row.capacity_slots,
                    row.populated_slots, row.capacity_bytes);
            }
        } else {
            let _ = writeln!(out, "prefill import-cache peak: no link snapshot recorded");
        }
        out
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn peak_survives_restoration_and_counts_cached_misses() {
        let mut imports: Vec<Vec<Option<Option<usize>>>> = Vec::with_capacity(4);
        let mut row = Vec::with_capacity(8);
        row.extend([None, Some(None), Some(Some(3))]);
        imports.push(row);
        let mut targets: Vec<Vec<Option<Option<u8>>>> = Vec::with_capacity(2);
        let mut row = Vec::with_capacity(4);
        row.push(Some(Some(7)));
        targets.push(row);
        let expected_import_bytes = imports.capacity() * std::mem::size_of_val(&imports[0])
            + imports[0].capacity() * std::mem::size_of::<Option<Option<usize>>>();
        let mut recorder = ImportCachePeakRecorder { enabled: true, ..Default::default() };
        assert!(recorder.record(&imports, &targets, 10, 12));
        let peak = recorder.peak.unwrap();
        assert_eq!(peak.imports.capacity_bytes, expected_import_bytes);
        assert_eq!(peak.imports.populated_slots, 2);
        assert_eq!(peak.call_targets.populated_slots, 1);
        assert_eq!((peak.first_chunk, peak.end_chunk), (10, 12));
        imports.clear();
        targets.clear();
        assert!(!recorder.record(&imports, &targets, 20, 21));
        assert_eq!(recorder.peak, Some(peak));
    }

    #[test]
    fn disabled_recorder_keeps_no_receipt() {
        let rows: Vec<Vec<Option<Option<usize>>>> = vec![vec![Some(Some(3))]];
        let mut recorder = ImportCachePeakRecorder::default();
        assert!(!recorder.record(&rows, &rows, 0, 1));
        assert!(recorder.peak.is_none());
    }
}
