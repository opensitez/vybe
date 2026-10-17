//! Private opt-in byte-indexed source map. Dense is the default. Mutable Vec
//! access explicitly materializes compact storage; read iteration stays sparse.
use std::{cell::Cell, fmt, ops::{Deref, DerefMut, Index}, sync::OnceLock};

thread_local! { static COMPACT_EMISSION: Cell<bool> = const { Cell::new(false) }; }
pub struct LineEmissionScope { previous: bool, _thread: std::marker::PhantomData<std::rc::Rc<()>> }
impl LineEmissionScope {
    pub fn enter(compact: bool) -> Self {
        Self { previous: COMPACT_EMISSION.with(|flag| flag.replace(compact)), _thread: std::marker::PhantomData }
    }
}
impl Drop for LineEmissionScope {
    fn drop(&mut self) { COMPACT_EMISSION.with(|flag| flag.set(self.previous)); }
}

#[derive(Clone)]
struct Run { end: usize, line: u32 }
#[derive(Clone)]
enum Storage {
    Dense(Vec<u32>),
    Runs { len: usize, runs: Vec<Run>, materialized: OnceLock<Vec<u32>> },
}
pub struct LineTable(Storage);
impl Clone for LineTable {
    fn clone(&self) -> Self {
        Self(match &self.0 {
            Storage::Dense(lines) => Storage::Dense(lines.clone()),
            // A borrowed Vec view is only a cache. Cloning it would duplicate
            // the full dense allocation during linked-code preparation.
            Storage::Runs { len, runs, .. } => Storage::Runs {
                len: *len, runs: runs.clone(), materialized: OnceLock::new(),
            },
        })
    }
}
impl Default for LineTable { fn default() -> Self { Self(Storage::Dense(Vec::new())) } }
impl From<Vec<u32>> for LineTable { fn from(lines: Vec<u32>) -> Self { Self(Storage::Dense(lines)) } }
impl LineTable {
    pub fn for_emission() -> Self {
        if COMPACT_EMISSION.with(Cell::get) {
            Self(Storage::Runs { len: 0, runs: Vec::new(), materialized: OnceLock::new() })
        } else { Self::default() }
    }
    pub fn len(&self) -> usize { match &self.0 { Storage::Dense(v) => v.len(), Storage::Runs {len,..} => *len } }
    pub fn is_empty(&self) -> bool { self.len() == 0 }
    // Compatibility capacity is in logical line entries, not physical bytes.
    // Use storage_bytes for allocation/retention measurements.
    pub fn capacity(&self) -> usize { match &self.0 {Storage::Dense(v)=>v.capacity(), Storage::Runs{len,..}=>*len} }
    pub fn storage_bytes(&self) -> usize {
        match &self.0 {
            Storage::Dense(v) => v.capacity() * std::mem::size_of::<u32>(),
            Storage::Runs{runs,materialized,..} => runs.capacity() * std::mem::size_of::<Run>()
                + materialized.get().map_or(0, |v| v.capacity()*std::mem::size_of::<u32>()),
        }
    }
    pub fn value_at(&self, offset: usize) -> Option<&u32> {
        match &self.0 {
            Storage::Dense(v) => v.get(offset),
            Storage::Runs {len,runs,..} if offset < *len =>
                runs.get(runs.partition_point(|run| run.end <= offset)).map(|run| &run.line),
            _ => None,
        }
    }
    fn append(&mut self, count: usize, line: u32) {
        if count == 0 {return;}
        match &mut self.0 {
            Storage::Dense(v) => v.resize(v.len().checked_add(count).expect("line offset overflow"), line),
            Storage::Runs {len,runs,materialized} => {
                *len = len.checked_add(count).expect("line offset overflow");
                match runs.last_mut() {
                    Some(run) if run.line == line => run.end = *len,
                    _ => runs.push(Run{end:*len,line}),
                }
                // A materialized borrowed Vec is no longer visible while we
                // hold &mut self. Invalidate it rather than serve old bytes.
                materialized.take();
            }
        }
        // Adversarial one-byte line changes must not retain a larger map than
        // the original representation. Small maps avoid eager fallback churn.
        let dense = matches!(&self.0, Storage::Runs{len,runs,..}
            if *len >= 64 && runs.capacity()*std::mem::size_of::<Run>() > *len*4);
        if dense { let _ = self.deref_mut(); }
    }
    pub fn push(&mut self, line: u32) {
        if let Storage::Dense(v)=&mut self.0 {v.push(line);} else {self.append(1,line);}
    }
    pub fn extend_from_slice(&mut self, values: &[u32]) {
        if let Storage::Dense(v)=&mut self.0 {v.extend_from_slice(values);return;}
        let mut at=0;
        while at<values.len() {
            let line=values[at];let begin=at;at+=1;
            while at<values.len() && values[at]==line {at+=1;}
            self.append(at-begin,line);
        }
    }
    /// The Chunk emitter already knows every byte has the same source line.
    /// Preserve the dense Vec operation; append one run without scanning N
    /// temporary line entries when compact emission is selected.
    pub fn extend_repeated<const N: usize>(&mut self, line: u32) {
        if let Storage::Dense(v) = &mut self.0 {
            v.extend_from_slice(&[line; N]);
        } else {
            self.append(N, line);
        }
    }
    pub fn enable_compact(&mut self) {
        if self.is_empty() { self.0=Storage::Runs{len:0,runs:Vec::new(),materialized:OnceLock::new()}; } else { self.compact(); }
    }
    /// Split a byte-indexed map without materializing sparse line entries.
    pub fn split_off_table(&mut self, at: usize) -> Self {
        assert!(at <= self.len(), "line split out of bounds");
        match &mut self.0 {
            Storage::Dense(lines) => lines.split_off(at).into(),
            Storage::Runs { len, runs, materialized } => {
                let old_len = *len;
                let split = runs.partition_point(|run| run.end <= at);
                let previous_end = split.checked_sub(1).map_or(0, |i| runs[i].end);
                let mut tail = runs.split_off(split);
                if at > previous_end {
                    runs.push(Run { end: at, line: tail[0].line });
                }
                for run in &mut tail { run.end -= at; }
                *len = at;
                materialized.take();
                Self(Storage::Runs { len: old_len - at, runs: tail, materialized: OnceLock::new() })
            }
        }
    }
    fn append_table(&mut self, other: Self) {
        if let Storage::Dense(lines) = &mut self.0 {
            lines.extend(Vec::<u32>::from(other));
            return;
        }
        match other.0 {
            Storage::Dense(lines) => self.extend_from_slice(&lines),
            Storage::Runs { runs, .. } => {
                let mut previous_end = 0;
                for run in runs {
                    self.append(run.end - previous_end, run.line);
                    previous_end = run.end;
                }
            }
        }
    }
    /// Replace an emitted instruction span, retaining compressed line runs.
    /// Dense callers keep the ordinary Vec splice path.
    pub fn replace_range_table(&mut self, range: std::ops::Range<usize>, replacement: Self) {
        assert!(range.start <= range.end && range.end <= self.len(), "line replacement out of bounds");
        if let Storage::Dense(lines) = &mut self.0 {
            lines.splice(range, Vec::<u32>::from(replacement));
            return;
        }
        let suffix = self.split_off_table(range.end);
        drop(self.split_off_table(range.start));
        self.append_table(replacement);
        self.append_table(suffix);
    }
    pub fn compact(&mut self) {
        if !matches!(self.0,Storage::Dense(_)) {return;}
        let dense=match std::mem::replace(&mut self.0,Storage::Dense(Vec::new())) {Storage::Dense(v)=>v,_=>unreachable!()};
        // The complete dense allocation is known. Never build a larger run
        // scratch buffer and only then discover that it cannot be profitable.
        let budget=dense.capacity()*std::mem::size_of::<u32>();
        let max_runs=budget/std::mem::size_of::<Run>();
        let mut runs:Vec<Run>=Vec::new();
        for (at,&line) in dense.iter().enumerate() {
            if let Some(run)=runs.last_mut().filter(|run|run.line==line) {
                run.end=at+1;
                continue;
            }
            if runs.len()==max_runs {
                self.0=Storage::Dense(dense);
                return;
            }
            if runs.len()==runs.capacity() {
                // Geometric growth stays amortized linear. The final growth
                // is capped rather than letting Vec double past the budget.
                let target=runs.capacity().saturating_mul(2).max(4).min(max_runs);
                runs.reserve_exact(target-runs.len());
                #[cfg(test)] COMPACTION_PEAK_RUN_BYTES.with(|peak|
                    peak.set(peak.get().max(runs.capacity()*std::mem::size_of::<Run>())));
            }
            runs.push(Run{end:at+1,line});
        }
        if runs.capacity()*std::mem::size_of::<Run>() < budget {
            self.0=Storage::Runs{len:dense.len(),runs,materialized:OnceLock::new()};
        } else {self.0=Storage::Dense(dense);}
    }
    pub fn iter(&self) -> LineIter<'_> {
        match &self.0 {
            Storage::Dense(v)=>LineIter(IterStorage::Dense(v.iter())),
            Storage::Runs{len,runs,..}=>LineIter(IterStorage::Runs{runs:runs.iter(),line:&ZERO,remaining:0,previous_end:0,left:*len}),
        }
    }
}
static ZERO:u32=0;
#[cfg(test)]
mod replacement_tests {
    use super::*;

    fn table(values: &[u32], sparse: bool) -> LineTable {
        let mut table = LineTable::from(values.to_vec());
        if sparse { table.enable_compact(); }
        table
    }

    #[test]
    fn split_and_replace_match_vec_at_every_run_boundary_and_interior() {
        let values: Vec<_> = (0..512).map(|i| (i / 64) as u32).collect();
        for sparse in [false, true] {
            for at in [0, 1, 63, 64, 65, 255, 511, 512] {
                let mut actual = table(&values, sparse);
                // Exercise invalidation of a previously materialized view.
                assert_eq!(&**actual, &values);
                let suffix = actual.split_off_table(at);
                assert_eq!(actual, values[..at].to_vec());
                assert_eq!(suffix, values[at..].to_vec());
                for start in [0, at / 2, at] {
                    for replacement_sparse in [false, true] {
                        for replacement in [vec![], vec![0; 128], vec![99; 192]] {
                            let mut expected = values[..at].to_vec();
                            expected.splice(start..at, replacement.clone());
                            let mut changed = actual.clone();
                            changed.replace_range_table(start..at, table(&replacement, replacement_sparse));
                            assert_eq!(changed, expected);
                            for i in 0..=expected.len() {
                                assert_eq!(changed.value_at(i), expected.get(i));
                            }
                        }
                    }
                }
            }
        }
    }

    #[test]
    fn large_helper_replacement_stays_sparse_without_dense_side_cache() {
        let mut source = table(&vec![7; 1_000_000], true);
        let appended = table(&vec![9; 10_000], true);
        source.append_table(appended);
        let replacement = source.split_off_table(1_000_000);
        source.replace_range_table(100..120, replacement);
        assert_eq!(source.len(), 1_009_980);
        assert_eq!(source.value_at(99), Some(&7));
        assert_eq!(source.value_at(100), Some(&9));
        assert_eq!(source.value_at(10_099), Some(&9));
        assert_eq!(source.value_at(10_100), Some(&7));
        assert!(matches!(&source.0, Storage::Runs { materialized, .. } if materialized.get().is_none()));
        assert!(source.storage_bytes() < 1024);
    }

    #[test]
    fn cloning_sparse_map_does_not_copy_a_dense_read_cache() {
        let source = table(&vec![17; 1_000_000], true);
        assert_eq!(source.as_slice().len(), 1_000_000);
        assert!(source.storage_bytes() >= 4_000_000);
        let mut cloned = source.clone();
        assert!(cloned.storage_bytes() < 1024);
        assert_eq!(source, cloned);
        cloned[3] = 22;
        assert_eq!(source.value_at(3), Some(&17));
        assert_eq!(cloned.value_at(3), Some(&22));
    }
}
#[cfg(test)]
mod uniform_append_tests {
    use super::*;
    fn compare<const N: usize>() {
        let _guard = LineEmissionScope::enter(true);
        let mut generic = LineTable::for_emission();
        let mut repeated = LineTable::for_emission();
        let mut dense = LineTable::default();
        for i in 0..2048 {
            let line = if i < 1024 { 0 } else { (i % 17) as u32 };
            generic.extend_from_slice(&[line; N]);
            repeated.extend_repeated::<N>(line);
            dense.extend_repeated::<N>(line);
            assert_eq!(generic, repeated);
            assert_eq!(dense, repeated);
            assert_eq!(generic.storage_bytes(), repeated.storage_bytes());
        }
        for at in 0..=generic.len() {
            assert_eq!(generic.value_at(at), repeated.value_at(at));
        }
    }
    #[test]
    fn uniform_emission_matches_generic_dense_and_sparse_including_fallback() {
        compare::<0>(); compare::<1>(); compare::<4>(); compare::<7>(); compare::<24>();
    }
}
pub struct LineIter<'a>(IterStorage<'a>);
enum IterStorage<'a> {
    Dense(std::slice::Iter<'a,u32>),
    Runs {runs:std::slice::Iter<'a,Run>,line:&'a u32,remaining:usize,previous_end:usize,left:usize},
}
impl<'a> Iterator for LineIter<'a> {
    type Item=&'a u32;
    fn next(&mut self)->Option<Self::Item> {
        match &mut self.0 {
            IterStorage::Dense(iter)=>iter.next(),
            IterStorage::Runs{runs,line,remaining,previous_end,left}=>{
                if *remaining==0 {let run=runs.next()?;*remaining=run.end-*previous_end;*previous_end=run.end;*line=&run.line;}
                *remaining-=1;*left-=1;Some(*line)
            }
        }
    }
    fn size_hint(&self)->(usize,Option<usize>){let n=match &self.0{IterStorage::Dense(v)=>v.len(),IterStorage::Runs{left,..}=>*left};(n,Some(n))}
}
impl ExactSizeIterator for LineIter<'_> {}
impl std::iter::FusedIterator for LineIter<'_> {}
impl<'a> IntoIterator for &'a LineTable {type Item=&'a u32;type IntoIter=LineIter<'a>;fn into_iter(self)->Self::IntoIter{self.iter()}}
impl fmt::Debug for LineTable {fn fmt(&self,f:&mut fmt::Formatter<'_>)->fmt::Result{f.debug_list().entries(self.iter()).finish()}}
impl PartialEq for LineTable {fn eq(&self,other:&Self)->bool{self.iter().eq(other.iter())}}
impl Eq for LineTable {}
impl PartialEq<Vec<u32>> for LineTable {fn eq(&self,other:&Vec<u32>)->bool{self.iter().eq(other.iter())}}
impl Deref for LineTable {
    type Target=Vec<u32>;
    fn deref(&self)->&Self::Target {
        match &self.0 {
            Storage::Dense(v)=>v,
            Storage::Runs{len,runs,materialized}=>materialized.get_or_init(||{
                let mut dense=Vec::with_capacity(*len);for run in runs{dense.resize(run.end,run.line);}dense
            }),
        }
    }
}
impl DerefMut for LineTable {
    fn deref_mut(&mut self)->&mut Self::Target {
        if matches!(self.0,Storage::Runs{..}) {
            let storage=std::mem::replace(&mut self.0,Storage::Dense(Vec::new()));
            let Storage::Runs{len,runs,materialized}=storage else{unreachable!()};
            let dense=materialized.into_inner().unwrap_or_else(||{let mut v=Vec::with_capacity(len);for run in runs{v.resize(run.end,run.line);}v});
            self.0=Storage::Dense(dense);
        }
        let Storage::Dense(v)=&mut self.0 else{unreachable!()};v
    }
}
impl std::ops::IndexMut<usize> for LineTable {fn index_mut(&mut self,at:usize)->&mut u32{&mut self.deref_mut()[at]}}
impl Index<usize> for LineTable {type Output=u32;fn index(&self,at:usize)->&u32{self.value_at(at).expect("line index out of bounds")}}

impl PartialEq<LineTable> for Vec<u32> {fn eq(&self,other:&LineTable)->bool{self.iter().eq(other.iter())}}
impl From<LineTable> for Vec<u32> {
    fn from(table:LineTable)->Self {
        match table.0 {
            Storage::Dense(v)=>v,
            Storage::Runs{len,runs,materialized}=>materialized.into_inner().unwrap_or_else(||{let mut v=Vec::with_capacity(len);for run in runs{v.resize(run.end,run.line);}v}),
        }
    }
}
impl IntoIterator for LineTable {type Item=u32;type IntoIter=std::vec::IntoIter<u32>;fn into_iter(self)->Self::IntoIter{Vec::<u32>::from(self).into_iter()}}
macro_rules! range_index {
    ($($range:ty),* $(,)?)=>{$(
        impl Index<$range> for LineTable {type Output=[u32];fn index(&self,range:$range)->&[u32]{&self.deref()[range]}}
        impl std::ops::IndexMut<$range> for LineTable {fn index_mut(&mut self,range:$range)->&mut[u32]{&mut self.deref_mut()[range]}}
    )*};
}
range_index!(std::ops::Range<usize>,std::ops::RangeFrom<usize>,std::ops::RangeTo<usize>,std::ops::RangeInclusive<usize>,std::ops::RangeToInclusive<usize>,std::ops::RangeFull);

#[cfg(test)] mod tests {
    use super::*;
    #[test] fn default_and_nested_scopes_restore_on_unwind(){
        assert!(matches!(LineTable::for_emission().0,Storage::Dense(_)));
        {let _guard=LineEmissionScope::enter(true);assert!(matches!(LineTable::for_emission().0,Storage::Runs{..}));
         {let _nested=LineEmissionScope::enter(false);assert!(matches!(LineTable::for_emission().0,Storage::Dense(_)));}
         assert!(matches!(LineTable::for_emission().0,Storage::Runs{..}));}
        let _=std::panic::catch_unwind(||{let _guard=LineEmissionScope::enter(true);panic!("scope");});
        assert!(matches!(LineTable::for_emission().0,Storage::Dense(_)));
    }
    #[test] fn logical_debug_iteration_and_offsets_match_vec(){
        let _guard=LineEmissionScope::enter(true);let mut compact=LineTable::for_emission();let mut dense=Vec::new();
        for n in 0..4096{let count=[1,4,7,8,12][n%5];let line=((n/20)%13) as u32;
            compact.extend_from_slice(&vec![line;count]);dense.extend_from_slice(&vec![line;count]);}
        assert_eq!(compact,dense);assert_eq!(format!("{compact:?}"),format!("{dense:?}"));
        for at in 0..dense.len()+1{assert_eq!(compact.value_at(at),dense.get(at));}
        assert_eq!(compact.iter().len(),dense.len());assert_eq!(compact.value_at(usize::MAX),None);
    }
    #[test] fn vec_mutation_fallback_clone_and_recompression(){
        let _guard=LineEmissionScope::enter(true);let mut compact=LineTable::for_emission();compact.extend_from_slice(&[7;100]);
        let old=compact.clone();let mut dense=vec![7;100];
        compact[0]=21;dense[0]=21;
        compact.splice(3..90,[1,2,3]);dense.splice(3..90,[1,2,3]);
        compact.truncate(9);dense.truncate(9);compact.resize(30,5);dense.resize(30,5);
        assert_eq!(compact,dense);assert_eq!(old,vec![7;100]);compact.compact();assert_eq!(compact,dense);
        assert_eq!(&compact[0..3],&dense[0..3]);assert_eq!(compact.get(0..3),dense.get(0..3));
        let tail=compact.split_off(7);let expected=dense.split_off(7);assert_eq!(tail,expected);assert_eq!(compact,dense);
    }
    #[test] fn append_invalidates_materialized_view(){
        let _guard=LineEmissionScope::enter(true);let mut table=LineTable::for_emission();table.extend_from_slice(&[4;100]);
        assert_eq!(table.as_slice(),&[4;100]);table.push(8);assert_eq!(table.as_slice()[100],8);
        assert_eq!(table.value_at(100),Some(&8));assert_eq!(table.iter().count(),101);
    }
    #[test] fn sparse_emission_and_adversarial_fallback(){
        let _guard=LineEmissionScope::enter(true);let mut table=LineTable::for_emission();
        for _ in 0..100_000{table.extend_from_slice(&[9;4]);}assert!(table.storage_bytes()<128);assert_eq!(table[399_999],9);
        let mut alternating=LineTable::for_emission();for n in 0..1024{alternating.push(n%2);}
        assert!(matches!(alternating.0,Storage::Dense(_)));assert_eq!(alternating.iter().copied().collect::<Vec<_>>(),(0..1024).map(|n|n%2).collect::<Vec<_>>());
    }
}

#[cfg(test)]
thread_local! { static COMPACTION_PEAK_RUN_BYTES: Cell<usize> = const { Cell::new(0) }; }
#[cfg(test)]
mod bounded_compaction_tests {
    use super::*;
    #[test]
    fn known_dense_budget_bounds_peak_scratch_and_preserves_every_byte() {
        for len in [0,1,3,4,15,16,17,63,64,65,1023,1024,1025,65_536] {
            for spare in [0,3,17] {
                for pattern in 0..3 {
                    let mut dense=Vec::with_capacity(len+spare);
                    dense.extend((0..len).map(|i|match pattern {0=>0,1=>(i%2)as u32,_=>(i/128)as u32}));
                    let budget=dense.capacity()*std::mem::size_of::<u32>();
                    let expected=dense.clone();
                    let mut table=LineTable::from(dense);
                    COMPACTION_PEAK_RUN_BYTES.with(|peak|peak.set(0));
                    table.compact();
                    assert!(COMPACTION_PEAK_RUN_BYTES.with(Cell::get)<=budget,
                        "scratch budget len={len} spare={spare} pattern={pattern}");
                    assert_eq!(table,expected);
                    assert_eq!(format!("{table:?}"),format!("{expected:?}"));
                    for at in 0..=len {assert_eq!(table.value_at(at),expected.get(at));}
                    table.push(u32::MAX);
                    assert_eq!(table.value_at(len),Some(&u32::MAX));
                    assert_eq!(table.value_at(len+1),None);
                }
            }
        }
    }
    #[test]
    fn long_uniform_tail_after_many_runs_stays_profitable_and_clone_is_exact() {
        let mut dense=vec![0;65_536];
        for (i,value) in dense.iter_mut().take(1024).enumerate(){*value=(i%2)as u32;}
        let mut table=LineTable::from(dense.clone());
        table.compact();
        assert!(table.storage_bytes()<dense.capacity()*4);
        let old=table.clone();
        table[32]=u32::MAX;
        assert_eq!(old,dense);
        assert_eq!(table.value_at(32),Some(&u32::MAX));
        assert_eq!(old.value_at(32),Some(&dense[32]));
    }
}
