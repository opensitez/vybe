//! LINQ: Chunk, MaxBy, MinBy, Append, Prepend, DefaultIfEmpty, Reverse, Flatten.
use super::helpers::run_csharp;

#[test]
fn chunk_splits_sequence_into_fixed_size_batches() {
    assert_eq!(
        run_csharp(
            r#"var batches=new[]{1,2,3,4,5}.Chunk(2).ToList();
Console.WriteLine(batches.Count);
Console.WriteLine(batches[0].Length);"#
        ),
        &["3", "2"]
    );
}

#[test]
fn max_by_returns_element_with_maximum_key() {
    assert_eq!(
        run_csharp(
            r#"var words=new[]{"a","bbb","cc"};
Console.WriteLine(words.MaxBy(w=>w.Length));"#
        ),
        &["bbb"]
    );
}

#[test]
fn min_by_returns_element_with_minimum_key() {
    assert_eq!(
        run_csharp(
            r#"var words=new[]{"a","bbb","cc"};
Console.WriteLine(words.MinBy(w=>w.Length));"#
        ),
        &["a"]
    );
}

#[test]
fn append_adds_element_to_end_of_sequence() {
    assert_eq!(
        run_csharp(
            r#"var result=new[]{1,2,3}.Append(4);
Console.WriteLine(result.Last());"#
        ),
        &["4"]
    );
}

#[test]
fn prepend_adds_element_to_start_of_sequence() {
    assert_eq!(
        run_csharp(
            r#"var result=new[]{2,3,4}.Prepend(1);
Console.WriteLine(result.First());"#
        ),
        &["1"]
    );
}

#[test]
fn default_if_empty_returns_default_for_empty_sequence() {
    assert_eq!(
        run_csharp(
            r#"var result=System.Array.Empty<int>().DefaultIfEmpty(99);
Console.WriteLine(result.First());"#
        ),
        &["99"]
    );
}

#[test]
fn reverse_inverts_order_of_elements() {
    assert_eq!(
        run_csharp(
            r#"var result=new[]{1,2,3}.Reverse();
foreach(var n in result) Console.WriteLine(n);"#
        ),
        &["3", "2", "1"]
    );
}

#[test]
fn element_at_returns_item_at_given_index() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(new[]{10,20,30}.ElementAt(1));"#),
        &["20"]
    );
}

#[test]
fn take_last_returns_trailing_elements() {
    assert_eq!(
        run_csharp(
            r#"var result=new[]{1,2,3,4,5}.TakeLast(2);
foreach(var n in result) Console.WriteLine(n);"#
        ),
        &["4", "5"]
    );
}

#[test]
fn skip_last_omits_trailing_elements() {
    assert_eq!(
        run_csharp(
            r#"var result=new[]{1,2,3,4,5}.SkipLast(2);
Console.WriteLine(result.Count());"#
        ),
        &["3"]
    );
}
