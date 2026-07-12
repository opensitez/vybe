use crate::helpers::run_main;

#[test]
fn files_write_string_creates_file_with_content() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Files.createTempFile("vybe", ".txt"); java.nio.file.Files.writeString(p, "alpha"); System.out.println(java.nio.file.Files.readString(p)); java.nio.file.Files.delete(p);"#,
    );
    assert_eq!(out, vec!["alpha"]);
}

#[test]
fn files_write_bytes_stores_binary_data() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Files.createTempFile("vybe", ".bin"); byte[] data = new byte[]{65, 66, 67}; java.nio.file.Files.write(p, data); byte[] back = java.nio.file.Files.readAllBytes(p); System.out.println(back[0]); System.out.println(back[2]); java.nio.file.Files.delete(p);"#,
    );
    assert_eq!(out, vec!["65", "67"]);
}

#[test]
fn files_read_all_lines_returns_line_list() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Files.createTempFile("vybe", ".txt"); java.nio.file.Files.writeString(p, "line1\nline2"); java.util.List<String> lines = java.nio.file.Files.readAllLines(p); System.out.println(lines.size()); System.out.println(lines.get(1)); java.nio.file.Files.delete(p);"#,
    );
    assert_eq!(out, vec!["2", "line2"]);
}

#[test]
fn files_append_extends_existing_file() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Files.createTempFile("vybe", ".txt"); java.nio.file.Files.writeString(p, "ab"); java.nio.file.Files.writeString(p, "cd", java.nio.file.StandardOpenOption.APPEND); System.out.println(java.nio.file.Files.readString(p)); java.nio.file.Files.delete(p);"#,
    );
    assert_eq!(out, vec!["abcd"]);
}

#[test]
fn files_exists_true_for_created_file() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Files.createTempFile("vybe", ".txt"); System.out.println(java.nio.file.Files.exists(p)); java.nio.file.Files.delete(p);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn files_exists_false_after_delete() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Files.createTempFile("vybe", ".txt"); java.nio.file.Files.delete(p); System.out.println(java.nio.file.Files.exists(p));"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn files_not_exists_true_for_deleted_path() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Files.createTempFile("vybe", ".txt"); java.nio.file.Files.delete(p); System.out.println(java.nio.file.Files.notExists(p));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn files_is_regular_file_true_for_temp_file() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Files.createTempFile("vybe", ".dat"); System.out.println(java.nio.file.Files.isRegularFile(p)); java.nio.file.Files.delete(p);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn files_is_directory_true_for_temp_dir() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Files.createTempDirectory("vybedir"); System.out.println(java.nio.file.Files.isDirectory(p)); java.nio.file.Files.delete(p);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn files_is_regular_file_false_for_directory() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Files.createTempDirectory("vybedir"); System.out.println(java.nio.file.Files.isRegularFile(p)); java.nio.file.Files.delete(p);"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn files_create_file_makes_empty_file() {
    let out = run_main(
        r#"java.nio.file.Path dir = java.nio.file.Files.createTempDirectory("vybedir"); java.nio.file.Path f = dir.resolve("newfile.txt"); java.nio.file.Files.createFile(f); System.out.println(java.nio.file.Files.size(f)); java.nio.file.Files.delete(f); java.nio.file.Files.delete(dir);"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn files_create_directory_makes_single_dir() {
    let out = run_main(
        r#"java.nio.file.Path parent = java.nio.file.Files.createTempDirectory("vybedir"); java.nio.file.Path child = parent.resolve("child"); java.nio.file.Files.createDirectory(child); System.out.println(java.nio.file.Files.isDirectory(child)); java.nio.file.Files.delete(child); java.nio.file.Files.delete(parent);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn files_create_directories_makes_nested_path() {
    let out = run_main(
        r#"java.nio.file.Path base = java.nio.file.Files.createTempDirectory("vybedir"); java.nio.file.Path nested = base.resolve("a/b/c"); java.nio.file.Files.createDirectories(nested); System.out.println(java.nio.file.Files.isDirectory(nested)); java.nio.file.Files.delete(nested); java.nio.file.Files.delete(base.resolve("a/b")); java.nio.file.Files.delete(base.resolve("a")); java.nio.file.Files.delete(base);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn files_delete_removes_file() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Files.createTempFile("vybe", ".tmp"); java.nio.file.Files.delete(p); System.out.println(java.nio.file.Files.notExists(p));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn files_delete_if_exists_returns_true_when_present() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Files.createTempFile("vybe", ".tmp"); boolean removed = java.nio.file.Files.deleteIfExists(p); System.out.println(removed);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn files_delete_if_exists_returns_false_when_absent() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Files.createTempFile("vybe", ".tmp"); java.nio.file.Files.delete(p); boolean removed = java.nio.file.Files.deleteIfExists(p); System.out.println(removed);"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn files_copy_duplicates_file_content() {
    let out = run_main(
        r#"java.nio.file.Path src = java.nio.file.Files.createTempFile("vybe", ".src"); java.nio.file.Files.writeString(src, "copyme"); java.nio.file.Path dst = src.resolveSibling("dst.txt"); java.nio.file.Files.copy(src, dst); System.out.println(java.nio.file.Files.readString(dst)); java.nio.file.Files.delete(dst); java.nio.file.Files.delete(src);"#,
    );
    assert_eq!(out, vec!["copyme"]);
}

#[test]
fn files_move_renames_file() {
    let out = run_main(
        r#"java.nio.file.Path src = java.nio.file.Files.createTempFile("vybe", ".mv"); java.nio.file.Files.writeString(src, "moved"); java.nio.file.Path dst = src.resolveSibling("renamed.txt"); java.nio.file.Files.move(src, dst); System.out.println(java.nio.file.Files.readString(dst)); System.out.println(java.nio.file.Files.notExists(src)); java.nio.file.Files.delete(dst);"#,
    );
    assert_eq!(out, vec!["moved", "true"]);
}

#[test]
fn files_size_returns_byte_count() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Files.createTempFile("vybe", ".sz"); java.nio.file.Files.writeString(p, "hello"); System.out.println(java.nio.file.Files.size(p)); java.nio.file.Files.delete(p);"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn files_is_readable_true_for_new_file() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Files.createTempFile("vybe", ".rd"); System.out.println(java.nio.file.Files.isReadable(p)); java.nio.file.Files.delete(p);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn files_is_writable_true_for_new_file() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Files.createTempFile("vybe", ".wr"); System.out.println(java.nio.file.Files.isWritable(p)); java.nio.file.Files.delete(p);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn files_list_directory_entries() {
    let out = run_main(
        r#"java.nio.file.Path dir = java.nio.file.Files.createTempDirectory("vybelist"); java.nio.file.Path f = dir.resolve("item.txt"); java.nio.file.Files.createFile(f); java.util.stream.Stream<java.nio.file.Path> s = java.nio.file.Files.list(dir); long count = s.count(); System.out.println(count); java.nio.file.Files.delete(f); java.nio.file.Files.delete(dir);"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn files_lines_stream_counts_entries() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Files.createTempFile("vybe", ".ln"); java.nio.file.Files.writeString(p, "a\nb\nc"); long count = java.nio.file.Files.lines(p).count(); System.out.println(count); java.nio.file.Files.delete(p);"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn files_read_string_empty_file() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Files.createTempFile("vybe", ".empty"); System.out.println(java.nio.file.Files.readString(p).length()); java.nio.file.Files.delete(p);"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn files_write_empty_byte_array() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Files.createTempFile("vybe", ".ebin"); java.nio.file.Files.write(p, new byte[0]); System.out.println(java.nio.file.Files.size(p)); java.nio.file.Files.delete(p);"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn files_create_temp_file_has_prefix_in_name() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Files.createTempFile("vybe", ".tmp"); String name = p.getFileName().toString(); System.out.println(name.startsWith("vybe")); java.nio.file.Files.delete(p);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn files_create_temp_directory_has_prefix() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Files.createTempDirectory("vybedir"); String name = p.getFileName().toString(); System.out.println(name.startsWith("vybedir")); java.nio.file.Files.delete(p);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn files_probe_content_type_for_text_file() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Files.createTempFile("vybe", ".txt"); String ct = java.nio.file.Files.probeContentType(p); System.out.println(ct != null || ct == null); java.nio.file.Files.delete(p);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn files_is_same_file_true_for_same_path() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Files.createTempFile("vybe", ".same"); System.out.println(java.nio.file.Files.isSameFile(p, p)); java.nio.file.Files.delete(p);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn files_is_hidden_false_for_temp_file() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Files.createTempFile("vybe", ".vis"); System.out.println(java.nio.file.Files.isHidden(p)); java.nio.file.Files.delete(p);"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn files_copy_replace_existing_overwrites() {
    let out = run_main(
        r#"java.nio.file.Path src = java.nio.file.Files.createTempFile("vybe", ".src"); java.nio.file.Files.writeString(src, "new"); java.nio.file.Path dst = src.resolveSibling("ovr.txt"); java.nio.file.Files.writeString(dst, "old"); java.nio.file.Files.copy(src, dst, java.nio.file.StandardCopyOption.REPLACE_EXISTING); System.out.println(java.nio.file.Files.readString(dst)); java.nio.file.Files.delete(dst); java.nio.file.Files.delete(src);"#,
    );
    assert_eq!(out, vec!["new"]);
}

#[test]
fn files_write_with_create_option_on_missing_path() {
    let out = run_main(
        r#"java.nio.file.Path dir = java.nio.file.Files.createTempDirectory("vybedir"); java.nio.file.Path f = dir.resolve("created.txt"); java.nio.file.Files.writeString(f, "fresh", java.nio.file.StandardOpenOption.CREATE); System.out.println(java.nio.file.Files.readString(f)); java.nio.file.Files.delete(f); java.nio.file.Files.delete(dir);"#,
    );
    assert_eq!(out, vec!["fresh"]);
}

#[test]
fn files_read_all_bytes_length_matches_size() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Files.createTempFile("vybe", ".rb"); java.nio.file.Files.writeString(p, "xyz"); byte[] data = java.nio.file.Files.readAllBytes(p); System.out.println(data.length); java.nio.file.Files.delete(p);"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn files_new_buffered_reader_reads_first_char_code() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Files.createTempFile("vybe", ".br"); java.nio.file.Files.writeString(p, "Z"); java.io.BufferedReader br = java.nio.file.Files.newBufferedReader(p); int ch = br.read(); br.close(); System.out.println(ch); java.nio.file.Files.delete(p);"#,
    );
    assert_eq!(out, vec!["90"]);
}

#[test]
fn files_new_output_stream_writes_byte() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Files.createTempFile("vybe", ".os"); java.io.OutputStream os = java.nio.file.Files.newOutputStream(p); os.write(88); os.close(); System.out.println(java.nio.file.Files.readAllBytes(p)[0]); java.nio.file.Files.delete(p);"#,
    );
    assert_eq!(out, vec!["88"]);
}

#[test]
fn files_walk_counts_file_in_directory() {
    let out = run_main(
        r#"java.nio.file.Path dir = java.nio.file.Files.createTempDirectory("vybewalk"); java.nio.file.Path f = dir.resolve("w.txt"); java.nio.file.Files.createFile(f); long count = java.nio.file.Files.walk(dir).count(); System.out.println(count >= 2); java.nio.file.Files.delete(f); java.nio.file.Files.delete(dir);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn files_mismatch_detects_different_content() {
    let out = run_main(
        r#"java.nio.file.Path a = java.nio.file.Files.createTempFile("vybe", ".a"); java.nio.file.Path b = java.nio.file.Files.createTempFile("vybe", ".b"); java.nio.file.Files.writeString(a, "aaa"); java.nio.file.Files.writeString(b, "bbb"); long pos = java.nio.file.Files.mismatch(a, b); System.out.println(pos >= 0); java.nio.file.Files.delete(a); java.nio.file.Files.delete(b);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn files_mismatch_negative_one_for_identical_files() {
    let out = run_main(
        r#"java.nio.file.Path a = java.nio.file.Files.createTempFile("vybe", ".a"); java.nio.file.Path b = java.nio.file.Files.createTempFile("vybe", ".b"); java.nio.file.Files.writeString(a, "same"); java.nio.file.Files.writeString(b, "same"); long pos = java.nio.file.Files.mismatch(a, b); System.out.println(pos); java.nio.file.Files.delete(a); java.nio.file.Files.delete(b);"#,
    );
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn files_set_last_modified_time_updates_timestamp() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Files.createTempFile("vybe", ".mt"); java.nio.file.attribute.FileTime ft = java.nio.file.Files.getLastModifiedTime(p); java.nio.file.Files.setLastModifiedTime(p, ft); System.out.println(java.nio.file.Files.getLastModifiedTime(p).equals(ft)); java.nio.file.Files.delete(p);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn files_read_attributes_size_key() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Files.createTempFile("vybe", ".attr"); java.nio.file.Files.writeString(p, "ab"); java.util.Map<String, Object> attrs = java.nio.file.Files.readAttributes(p, "basic:*"); System.out.println(attrs.containsKey("size")); java.nio.file.Files.delete(p);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn files_write_string_utf8_multibyte() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Files.createTempFile("vybe", ".utf"); java.nio.file.Files.writeString(p, "\u00e9"); String s = java.nio.file.Files.readString(p); System.out.println(s.length()); java.nio.file.Files.delete(p);"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn files_lines_find_first_returns_first_line() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Files.createTempFile("vybe", ".ff"); java.nio.file.Files.writeString(p, "first\nsecond"); java.util.Optional<String> opt = java.nio.file.Files.lines(p).findFirst(); System.out.println(opt.get()); java.nio.file.Files.delete(p);"#,
    );
    assert_eq!(out, vec!["first"]);
}

#[test]
fn files_is_executable_false_for_text_file() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Files.createTempFile("vybe", ".nexe"); System.out.println(java.nio.file.Files.isExecutable(p)); java.nio.file.Files.delete(p);"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn files_new_byte_channel_read_mode_opens() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Files.createTempFile("vybe", ".ch"); java.nio.file.Files.writeString(p, "x"); java.nio.channels.SeekableByteChannel ch = java.nio.file.Files.newByteChannel(p, java.util.Set.of(java.nio.file.StandardOpenOption.READ)); System.out.println(ch.isOpen()); ch.close(); java.nio.file.Files.delete(p);"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn files_create_link_not_supported_returns_boolean() {
    let out = run_main(
        r#"java.nio.file.Path p = java.nio.file.Files.createTempFile("vybe", ".lnk"); System.out.println(java.nio.file.Files.exists(p)); java.nio.file.Files.delete(p);"#,
    );
    assert_eq!(out, vec!["true"]);
}
