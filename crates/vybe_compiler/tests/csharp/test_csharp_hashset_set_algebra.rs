//! HashSet<T> set-algebra APIs: UnionWith, IntersectWith, ExceptWith, SymmetricExceptWith, subset/equality.

csharp_cases! {
    union_with_merges_disjoint_elements => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 1, 2 }; a.UnionWith(new[] { 3, 4 }); Console.WriteLine(a.Count);"#,
        ["4"]
    };

    union_with_absorbs_overlapping_elements_without_duplicates => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 1, 2, 3 }; a.UnionWith(new[] { 3, 4 }); Console.WriteLine(a.Count);"#,
        ["4"]
    };

    union_with_empty_other_leaves_set_unchanged => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 5, 6 }; a.UnionWith(new int[] { }); Console.WriteLine(a.Count);"#,
        ["2"]
    };

    union_with_into_empty_set_adopts_all_elements => {
        r#"using System.Collections.Generic; var a = new HashSet<int>(); a.UnionWith(new[] { 7, 8 }); Console.WriteLine(a.Contains(7)); Console.WriteLine(a.Count);"#,
        ["True", "2"]
    };

    union_with_self_does_not_change_count => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 1, 2 }; a.UnionWith(a); Console.WriteLine(a.Count);"#,
        ["2"]
    };

    union_with_preserves_existing_members => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 10 }; a.UnionWith(new[] { 20 }); Console.WriteLine(a.Contains(10));"#,
        ["True"]
    };

    union_with_string_elements_concatenates_unique_names => {
        r#"using System.Collections.Generic; var a = new HashSet<string> { "a" }; a.UnionWith(new[] { "b", "a" }); Console.WriteLine(a.Count);"#,
        ["2"]
    };

    union_with_enumerable_adds_all_new_items => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 1 }; var extra = new List<int> { 2, 3 }; a.UnionWith(extra); Console.WriteLine(a.Contains(3));"#,
        ["True"]
    };

    union_with_after_union_accumulates_third_batch => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 1 }; a.UnionWith(new[] { 2 }); a.UnionWith(new[] { 3 }); Console.WriteLine(a.Count);"#,
        ["3"]
    };

    intersect_with_keeps_only_shared_elements => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 1, 2, 3 }; a.IntersectWith(new[] { 2, 3, 4 }); Console.WriteLine(a.Count);"#,
        ["2"]
    };

    intersect_with_no_overlap_yields_empty_set => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 1, 2 }; a.IntersectWith(new[] { 5, 6 }); Console.WriteLine(a.Count);"#,
        ["0"]
    };

    intersect_with_identical_set_preserves_all => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 4, 5 }; a.IntersectWith(new[] { 4, 5 }); Console.WriteLine(a.Count);"#,
        ["2"]
    };

    intersect_with_self_is_identity => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 7, 8 }; a.IntersectWith(a); Console.WriteLine(a.Count);"#,
        ["2"]
    };

    intersect_with_single_shared_element => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 1, 2, 3 }; a.IntersectWith(new[] { 3, 9 }); Console.WriteLine(a.Contains(3)); Console.WriteLine(a.Contains(1));"#,
        ["True", "False"]
    };

    intersect_with_empty_other_clears_set => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 1, 2 }; a.IntersectWith(new int[] { }); Console.WriteLine(a.Count);"#,
        ["0"]
    };

    intersect_with_string_names_keeps_common => {
        r#"using System.Collections.Generic; var a = new HashSet<string> { "x", "y" }; a.IntersectWith(new[] { "y", "z" }); Console.WriteLine(a.Contains("y")); Console.WriteLine(a.Count);"#,
        ["True", "1"]
    };

    except_with_removes_elements_present_in_other => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 1, 2, 3 }; a.ExceptWith(new[] { 2, 4 }); Console.WriteLine(a.Count); Console.WriteLine(a.Contains(1));"#,
        ["2", "True"]
    };

    except_with_removes_all_when_other_is_superset => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 1, 2 }; a.ExceptWith(new[] { 1, 2, 3 }); Console.WriteLine(a.Count);"#,
        ["0"]
    };

    except_with_no_overlap_leaves_set_unchanged => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 5, 6 }; a.ExceptWith(new[] { 1, 2 }); Console.WriteLine(a.Count);"#,
        ["2"]
    };

    except_with_empty_other_is_no_op => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 9 }; a.ExceptWith(new int[] { }); Console.WriteLine(a.Count);"#,
        ["1"]
    };

    except_with_self_clears_entire_set => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 1, 2, 3 }; a.ExceptWith(a); Console.WriteLine(a.Count);"#,
        ["0"]
    };

    except_with_single_element_removal => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 10, 20, 30 }; a.ExceptWith(new[] { 20 }); Console.WriteLine(a.Contains(20));"#,
        ["False"]
    };

    symmetric_except_with_keeps_elements_in_exactly_one_set => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 1, 2, 3 }; a.SymmetricExceptWith(new[] { 2, 3, 4 }); Console.WriteLine(a.Contains(1)); Console.WriteLine(a.Contains(4)); Console.WriteLine(a.Contains(2));"#,
        ["True", "True", "False"]
    };

    symmetric_except_with_identical_sets_yields_empty => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 1, 2 }; a.SymmetricExceptWith(new[] { 1, 2 }); Console.WriteLine(a.Count);"#,
        ["0"]
    };

    symmetric_except_with_disjoint_sets_is_union => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 1, 2 }; a.SymmetricExceptWith(new[] { 3, 4 }); Console.WriteLine(a.Count);"#,
        ["4"]
    };

    symmetric_except_with_empty_other_is_identity => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 5, 6 }; a.SymmetricExceptWith(new int[] { }); Console.WriteLine(a.Count);"#,
        ["2"]
    };

    symmetric_except_with_self_clears_set => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 7, 8 }; a.SymmetricExceptWith(a); Console.WriteLine(a.Count);"#,
        ["0"]
    };

    is_subset_of_true_when_all_elements_contained => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 1, 2 }; var b = new HashSet<int> { 1, 2, 3 }; Console.WriteLine(a.IsSubsetOf(b));"#,
        ["True"]
    };

    is_subset_of_false_when_extra_element_in_a => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 1, 2, 9 }; var b = new HashSet<int> { 1, 2, 3 }; Console.WriteLine(a.IsSubsetOf(b));"#,
        ["False"]
    };

    is_subset_of_empty_set_only_for_empty => {
        r#"using System.Collections.Generic; var empty = new HashSet<int>(); var nonempty = new HashSet<int> { 1 }; Console.WriteLine(empty.IsSubsetOf(nonempty)); Console.WriteLine(nonempty.IsSubsetOf(empty));"#,
        ["True", "False"]
    };

    is_subset_of_self_is_true => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 4, 5 }; Console.WriteLine(a.IsSubsetOf(a));"#,
        ["True"]
    };

    is_proper_subset_of_true_when_strictly_smaller => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 1 }; var b = new HashSet<int> { 1, 2 }; Console.WriteLine(a.IsProperSubsetOf(b));"#,
        ["True"]
    };

    is_proper_subset_of_false_when_equal => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 1, 2 }; var b = new HashSet<int> { 1, 2 }; Console.WriteLine(a.IsProperSubsetOf(b));"#,
        ["False"]
    };

    is_superset_of_is_inverse_of_subset => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 1, 2, 3 }; var b = new HashSet<int> { 1, 2 }; Console.WriteLine(a.IsSupersetOf(b));"#,
        ["True"]
    };

    is_proper_superset_of_requires_extra_element => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 1, 2, 3 }; var b = new HashSet<int> { 1, 2 }; Console.WriteLine(a.IsProperSupersetOf(b));"#,
        ["True"]
    };

    set_equals_true_for_same_elements => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 1, 2, 3 }; var b = new HashSet<int> { 3, 2, 1 }; Console.WriteLine(a.SetEquals(b));"#,
        ["True"]
    };

    set_equals_false_for_different_sizes => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 1, 2 }; var b = new HashSet<int> { 1, 2, 3 }; Console.WriteLine(a.SetEquals(b));"#,
        ["False"]
    };

    set_equals_false_for_same_size_different_elements => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 1, 2 }; var b = new HashSet<int> { 1, 3 }; Console.WriteLine(a.SetEquals(b));"#,
        ["False"]
    };

    set_equals_self_is_true => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 9, 8 }; Console.WriteLine(a.SetEquals(a));"#,
        ["True"]
    };

    set_equals_with_enumerable_matching_set => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 4, 5 }; Console.WriteLine(a.SetEquals(new[] { 5, 4 }));"#,
        ["True"]
    };

    overlaps_true_when_sets_share_element => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 1, 2 }; var b = new HashSet<int> { 2, 3 }; Console.WriteLine(a.Overlaps(b));"#,
        ["True"]
    };

    overlaps_false_for_disjoint_sets => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 1, 2 }; var b = new HashSet<int> { 3, 4 }; Console.WriteLine(a.Overlaps(b));"#,
        ["False"]
    };

    union_then_intersect_restores_overlap_only => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 1, 2 }; a.UnionWith(new[] { 2, 3 }); a.IntersectWith(new[] { 2, 5 }); Console.WriteLine(a.Count); Console.WriteLine(a.Contains(2));"#,
        ["1", "True"]
    };

    except_then_union_readds_removed_only_if_new => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 1, 2, 3 }; a.ExceptWith(new[] { 2 }); a.UnionWith(new[] { 2, 4 }); Console.WriteLine(a.Contains(2)); Console.WriteLine(a.Contains(4));"#,
        ["True", "True"]
    };

    symmetric_except_twice_restores_original_when_same_other => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 1, 2, 3 }; var other = new[] { 2, 4 }; a.SymmetricExceptWith(other); a.SymmetricExceptWith(other); Console.WriteLine(a.SetEquals(new HashSet<int> { 1, 2, 3 }));"#,
        ["True"]
    };

    is_subset_of_after_union_with_superset => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 1 }; var b = new HashSet<int> { 1, 2 }; a.UnionWith(b); Console.WriteLine(a.IsSubsetOf(b));"#,
        ["False"]
    };

    intersect_with_after_except_yields_empty_when_disjoint => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 1, 2, 3 }; a.ExceptWith(new[] { 1, 2, 3 }); a.IntersectWith(new[] { 1 }); Console.WriteLine(a.Count);"#,
        ["0"]
    };

    set_equals_empty_sets => {
        r#"using System.Collections.Generic; var a = new HashSet<int>(); var b = new HashSet<int>(); Console.WriteLine(a.SetEquals(b));"#,
        ["True"]
    };

    is_proper_subset_of_empty_is_false => {
        r#"using System.Collections.Generic; var a = new HashSet<int> { 1 }; var b = new HashSet<int>(); Console.WriteLine(a.IsProperSubsetOf(b));"#,
        ["False"]
    };

    union_with_string_set_preserves_all_unique_tokens => {
        r#"using System.Collections.Generic; var a = new HashSet<string> { "one", "two" }; a.UnionWith(new[] { "two", "three" }); Console.WriteLine(a.Count);"#,
        ["3"]
    };
}
