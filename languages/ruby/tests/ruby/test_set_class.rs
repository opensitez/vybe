//! Real `Set` / `SortedSet` — a deduped tagged-array collection. `SortedSet`
//! keeps ascending order via the shared sorted core. (No hardcoded answers.)

macro_rules! ruby_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(super::helpers::run_ruby_one($src), $expected);
        }
    };
}

ruby_test!(
    set_add_dedupes_size,
    "require 'set'; s = Set.new; s.add(1); s.add(2); s.add(2); puts s.size",
    "2"
);

ruby_test!(
    set_include_true,
    "require 'set'; s = Set.new; s.add(3); puts s.include?(3)",
    "true"
);

ruby_test!(
    set_include_false,
    "require 'set'; s = Set.new; s.add(3); puts s.include?(9)",
    "false"
);

ruby_test!(
    set_to_a_sorted,
    "require 'set'; s = Set.new; s.add(3); s.add(1); s.add(2); puts s.to_a.sort.join('-')",
    "1-2-3"
);

ruby_test!(
    set_new_from_array_dedupes,
    "require 'set'; s = Set.new([1, 2, 2, 3, 1]); puts s.size",
    "3"
);

ruby_test!(
    set_add_returns_self_size,
    "require 'set'; s = Set.new; s.add(5).add(6); puts s.size",
    "2"
);

ruby_test!(
    sorted_set_iterates_ascending,
    "require 'set'; s = SortedSet.new; s.add(5); s.add(1); s.add(3); puts s.to_a.join('-')",
    "1-3-5"
);

ruby_test!(
    sorted_set_from_array_is_sorted,
    "require 'set'; s = SortedSet.new([30, 10, 20]); puts s.to_a.join('-')",
    "10-20-30"
);

ruby_test!(
    set_delete_removes_element,
    "require 'set'; s = Set.new([1, 2, 3]); s.delete(2); puts s.to_a.sort.join('-')",
    "1-3"
);

ruby_test!(
    set_clear_then_empty,
    "require 'set'; s = Set.new([1, 2]); s.clear; puts s.empty?",
    "true"
);

ruby_test!(
    set_merge_unions_dedup,
    "require 'set'; s = Set.new([1]); s.merge([2, 1, 3]); puts s.to_a.sort.join('-')",
    "1-2-3"
);

ruby_test!(
    set_replace_swaps_contents,
    "require 'set'; s = Set.new([1]); s.replace([2, 3]); puts s.to_a.sort.join('-')",
    "2-3"
);

ruby_test!(
    sorted_set_merge_stays_sorted,
    "require 'set'; s = SortedSet.new([5]); s.merge([3, 8, 1]); puts s.to_a.join('-')",
    "1-3-5-8"
);

// The exact programs that were previously hardcoded in the walker — now they
// run through the real Set class and must produce the same answers.
ruby_test!(
    set_dewaked_new_to_a_sort,
    "require 'set'; s = Set.new([1, 2, 2, 3]); puts s.to_a.sort.join('-')",
    "1-2-3"
);

ruby_test!(
    set_dewaked_add_size,
    "require 'set'; s = Set.new; s.add(1); s.add(1); puts s.size",
    "1"
);

ruby_test!(
    set_dewaked_delete,
    "require 'set'; s = Set.new([1, 2]); s.delete(1); puts s.to_a.join('-')",
    "2"
);

ruby_test!(
    set_dewaked_include,
    "require 'set'; s = Set.new([1]); puts s.include?(1)",
    "true"
);

ruby_test!(
    set_dewaked_clear_empty,
    "require 'set'; s = Set.new([1]); s.clear; puts s.empty?",
    "true"
);

ruby_test!(
    set_dewaked_replace,
    "require 'set'; s = Set.new([1]); s.replace([2, 3]); puts s.to_a.sort.join('-')",
    "2-3"
);

ruby_test!(
    set_dewaked_merge,
    "require 'set'; s = Set.new([1]); s.merge([2, 1]); puts s.to_a.sort.join('-')",
    "1-2"
);
