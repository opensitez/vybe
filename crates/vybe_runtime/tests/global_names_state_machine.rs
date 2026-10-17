//! Standalone std-only differential gate for persistent global-name metadata.
//! Also runnable directly with rustc --test, without building the PHP compiler.
#[path = "../src/segmented_global_names.rs"]
mod segmented_global_names;
#[path = "../src/global_name_index.rs"]
mod global_name_index;

use global_name_index::GlobalNameIndex;
use segmented_global_names::SegmentedGlobalNames;

fn reference(index: &GlobalNameIndex, slots: usize) -> Vec<String> {
    let mut names = vec![String::new(); slots];
    // An explicit deterministic alias policy. The cache must preserve whatever
    // exact policy its authoritative builder supplies, not choose its own.
    let mut entries: Vec<_> = index.iter().collect();
    entries.sort_by(|a, b| a.0.cmp(b.0));
    for (name, slot) in entries {
        if let Some(target) = names.get_mut(*slot as usize) { *target = name.clone(); }
    }
    names
}

fn snapshot(index: &GlobalNameIndex, slots: usize) -> SegmentedGlobalNames {
    index.cached_segmented_table(slots, || reference(index, slots))
}

#[test]
fn randomized_mutations_match_a_full_rebuild_and_keep_old_views_immutable() {
    let mut random = 0xd1ff_e7e1_1a5e_u64;
    let mut index = GlobalNameIndex::default();
    let mut slots = 0usize;
    let mut retained = Vec::new();
    for step in 0..10_000 {
        random ^= random << 13;
        random ^= random >> 7;
        random ^= random << 17;
        let name = format!("namespace.name_{}", (random >> 8) % 64);
        let slot = ((random >> 20) % 80) as u32;
        match random % 8 {
            0 => { index.insert(name, slot); }
            1 => { index.remove(&name); }
            2 => { *index.entry(name).or_insert(slot) = slot; }
            3 => { slots = (random >> 32) as usize % 96; }
            4 => {
                for (_, value) in &mut index {
                    *value = value.wrapping_add(1) % 80;
                }
            }
            5 => {
                index.insert(format!("trailing_{step}"), slots as u32);
                slots += 1;
            }
            6 => {
                let mut fork = index.clone();
                fork.insert("fork_only".into(), slot);
                let expected = reference(&fork, slots);
                assert_eq!(snapshot(&fork, slots).iter().cloned().collect::<Vec<_>>(), expected);
            }
            _ => {
                if random & 0x100 != 0 { index.clear(); }
            }
        }
        let expected = reference(&index, slots);
        let actual = snapshot(&index, slots);
        assert_eq!(actual.len(), expected.len(), "length at step {step}");
        assert_eq!(actual.iter().cloned().collect::<Vec<_>>(), expected, "names at step {step}");
        if step % 500 == 0 { retained.push((actual, expected)); }
    }
    for (view, expected) in retained {
        assert_eq!(view.iter().cloned().collect::<Vec<_>>(), expected);
    }
}
