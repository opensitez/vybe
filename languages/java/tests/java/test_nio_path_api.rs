use crate::helpers::run_main;

#[test]
fn path_get_name_returns_final_segment() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Paths.get("/home/user/doc.txt"); System.out.println(p.getFileName());"#,
    );
    assert_eq!(out, vec!["doc.txt"]);
}

#[test]
fn path_get_parent_returns_containing_directory() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Paths.get("/a/b/c"); System.out.println(p.getParent());"#,
    );
    assert_eq!(out, vec!["/a/b"]);
}

#[test]
fn path_get_root_returns_root_on_absolute_unix() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Paths.get("/var/log"); System.out.println(p.getRoot());"#,
    );
    assert_eq!(out, vec!["/"]);
}

#[test]
fn path_get_name_count_returns_segment_count() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Paths.get("a/b/c"); System.out.println(p.getNameCount());"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn path_get_index_zero_is_first_segment() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Paths.get("x/y/z"); System.out.println(p.getName(0));"#,
    );
    assert_eq!(out, vec!["x"]);
}

#[test]
fn path_get_index_last_is_final_segment() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Paths.get("x/y/z"); System.out.println(p.getName(2));"#,
    );
    assert_eq!(out, vec!["z"]);
}

#[test]
fn path_resolve_appends_relative_segment() {
    let out = run_main(
        r#"java.nio.file.Path base = java.nio.file.Paths.get("/base"); java.nio.file.Path p = base.resolve("child"); System.out.println(p.toString());"#,
    );
    assert_eq!(out, vec!["/base/child"]);
}

#[test]
fn path_resolve_absolute_replaces_base() {
    let out = run_main(
        r#"java.nio.file.Path base = java.nio.file.Paths.get("/base"); java.nio.file.Path p = base.resolve("/abs"); System.out.println(p.toString());"#,
    );
    assert_eq!(out, vec!["/abs"]);
}

#[test]
fn path_resolve_sibling_next_to_parent() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Paths.get("/a/b/file.txt"); System.out.println(p.resolveSibling("other.txt"));"#,
    );
    assert_eq!(out, vec!["/a/b/other.txt"]);
}

#[test]
fn path_relativize_strips_common_prefix() {
    let out = run_main(
        r#"java.nio.file.Path a = java.nio.file.Paths.get("/a/b/c"); java.nio.file.Path b = java.nio.file.Paths.get("/a/b/d"); System.out.println(a.relativize(b));"#,
    );
    assert_eq!(out, vec!["../d"]);
}

#[test]
fn path_normalize_collapses_dot_dot() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Paths.get("/a/b/../c"); System.out.println(p.normalize());"#,
    );
    assert_eq!(out, vec!["/a/c"]);
}

#[test]
fn path_normalize_removes_dot_segment() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Paths.get("/a/./b"); System.out.println(p.normalize());"#,
    );
    assert_eq!(out, vec!["/a/b"]);
}

#[test]
fn path_starts_with_prefix_true() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Paths.get("/usr/local/bin"); System.out.println(p.startsWith("/usr"));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn path_starts_with_prefix_false() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Paths.get("/usr/local"); System.out.println(p.startsWith("/var"));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn path_ends_with_segment_true() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Paths.get("/data/archive.zip"); System.out.println(p.endsWith("archive.zip"));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn path_ends_with_segment_false() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Paths.get("/data/file.txt"); System.out.println(p.endsWith(".zip"));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn path_is_absolute_true_for_rooted_path() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Paths.get("/abs/path"); System.out.println(p.isAbsolute());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn path_is_absolute_false_for_relative() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Paths.get("rel/path"); System.out.println(p.isAbsolute());"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn path_to_absolute_path_on_relative() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Paths.get("rel"); System.out.println(p.toAbsolutePath().isAbsolute());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn path_to_uri_roundtrip_scheme() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Paths.get("/tmp/test"); java.net.URI u = p.toUri(); System.out.println(u.getScheme());"#,
    );
    assert_eq!(out, vec!["file"]);
}

#[test]
fn path_from_uri_reconstructs_path() {
    let out = run_main(
        r#"java.net.URI u = java.nio.file.Paths.get("/tmp/fromuri").toUri(); java.nio.file.Path p = java.nio.file.Paths.get(u); System.out.println(p.getFileName());"#,
    );
    assert_eq!(out, vec!["fromuri"]);
}

#[test]
fn path_compare_to_lexicographic_order() {
    let out = run_main(
        r#"java.nio.file.Path a = java.nio.file.Paths.get("/a"); java.nio.file.Path b = java.nio.file.Paths.get("/b"); System.out.println(a.compareTo(b) < 0);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn path_equals_same_path_true() {
    let out = run_main(
        r#"java.nio.file.Path a = java.nio.file.Paths.get("/eq/path"); java.nio.file.Path b = java.nio.file.Paths.get("/eq/path"); System.out.println(a.equals(b));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn path_equals_different_path_false() {
    let out = run_main(
        r#"java.nio.file.Path a = java.nio.file.Paths.get("/eq/a"); java.nio.file.Path b = java.nio.file.Paths.get("/eq/b"); System.out.println(a.equals(b));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn path_hash_code_equal_for_equal_paths() {
    let out = run_main(
        r#"java.nio.file.Path a = java.nio.file.Paths.get("/hc/x"); java.nio.file.Path b = java.nio.file.Paths.get("/hc/x"); System.out.println(a.hashCode() == b.hashCode());"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn path_to_string_returns_path_text() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Paths.get("a", "b", "c"); System.out.println(p.toString());"#,
    );
    assert_eq!(out, vec!["a/b/c"]);
}

#[test]
fn paths_get_varargs_joins_segments() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Paths.get("home", "user", "docs"); System.out.println(p.getNameCount());"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn path_subpath_extracts_range() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Paths.get("a/b/c/d"); System.out.println(p.subpath(1, 3));"#,
    );
    assert_eq!(out, vec!["b/c"]);
}

#[test]
fn path_iterator_counts_segments() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Paths.get("one/two/three"); int n = 0; for (java.nio.file.Path seg : p) { n++; } System.out.println(n);"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn path_get_parent_null_for_single_name() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Paths.get("solo"); System.out.println(p.getParent() == null);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn path_resolve_multiple_segments() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Paths.get("/root").resolve("a").resolve("b"); System.out.println(p.toString());"#,
    );
    assert_eq!(out, vec!["/root/a/b"]);
}

#[test]
fn path_starts_with_full_path() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Paths.get("/a/b/c"); System.out.println(p.startsWith("/a/b/c"));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn path_ends_with_multi_segment_suffix() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Paths.get("/x/y/z"); System.out.println(p.endsWith("y/z"));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn path_normalize_idempotent() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Paths.get("/a/../b"); System.out.println(p.normalize().equals(p.normalize()));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn path_empty_relative_has_zero_name_count() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Paths.get(""); System.out.println(p.getNameCount());"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn path_file_name_on_root_only_is_root() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Paths.get("/"); System.out.println(p.getFileName());"#,
    );
    assert_eq!(out, vec!["/"]);
}

#[test]
fn path_relativize_same_path_is_empty() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Paths.get("/same/x"); System.out.println(p.relativize(p).toString());"#,
    );
    assert_eq!(out, vec!["."]);
}

#[test]
fn path_to_uri_contains_file_scheme() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Paths.get("/uri/test"); System.out.println(p.toUri().getScheme());"#,
    );
    assert_eq!(out, vec!["file"]);
}

#[test]
fn path_resolve_string_segment() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Paths.get("/base").resolve("file.dat"); System.out.println(p.getFileName());"#,
    );
    assert_eq!(out, vec!["file.dat"]);
}

#[test]
fn path_get_root_null_on_relative() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Paths.get("no/root"); System.out.println(p.getRoot() == null);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn path_compare_to_zero_for_equal() {
    let out = run_main(
        r#"java.nio.file.Path a = java.nio.file.Paths.get("/cmp/same"); java.nio.file.Path b = java.nio.file.Paths.get("/cmp/same"); System.out.println(a.compareTo(b));"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn path_resolve_sibling_at_root_level() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Paths.get("/only.txt"); System.out.println(p.resolveSibling("sibling.txt").getFileName());"#,
    );
    assert_eq!(out, vec!["sibling.txt"]);
}
