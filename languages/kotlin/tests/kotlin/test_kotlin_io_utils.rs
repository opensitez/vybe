use crate::helpers::run_prints;

#[test]
fn test_kotlin_io_writes_and_reads_text() {
    let out = run_prints(r#"
        fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_text_" + System.nanoTime() + "_a.txt")
            file.writeText("hello")
            println(file.exists())
            println(file.readText())
            file.delete()
        }
    "#);
    assert_eq!(out, &[
        "true",
        "hello"
    ]);
}

#[test]
fn test_kotlin_io_append_text_extends_existing_content() {
    let out = run_prints(r#"
        fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_append_" + System.nanoTime() + ".txt")
            file.writeText("a")
            file.appendText("b")
            file.appendText("c")
            println(file.readText())
            file.delete()
        }
    "#);
    assert_eq!(out, &["abc"]);
}

#[test]
fn test_kotlin_io_write_text_overwrites_previous_content() {
    let out = run_prints(r#"
        fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_overwrite_" + System.nanoTime() + ".txt")
            file.writeText("first")
            file.writeText("second")
            println(file.readText())
            file.delete()
        }
    "#);
    assert_eq!(out, &["second"]);
}

#[test]
fn test_kotlin_io_read_lines_preserves_blank_entries() {
    let out = run_prints(r#"
        fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_lines_" + System.nanoTime() + ".txt")
            file.writeText("a\n\n b\n")
            val lines = file.readLines()
            println(lines.size)
            println(lines.joinToString("|"))
            file.delete()
        }
    "#);
    assert_eq!(out, &[
        "3",
        "a|| b"
    ]);
}

#[test]
fn test_kotlin_io_for_each_line_collects_each_line() {
    let out = run_prints(r#"
        fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_foreach_" + System.nanoTime() + ".txt")
            file.writeText("u\nv\nw")
            val joined = StringBuilder()
            file.forEachLine { joined.append(it).append(".") }
            println(joined.toString())
            file.delete()
        }
    "#);
    assert_eq!(out, &["u.v.w."]);
}

#[test]
fn test_kotlin_io_use_lines_counts_lines() {
    let out = run_prints(r#"
        fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_use_lines_" + System.nanoTime() + ".txt")
            file.writeText("1\n2\n3")
            val count = file.useLines { lines -> lines.count() }
            println(count)
            file.delete()
        }
    "#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_kotlin_io_write_and_read_bytes() {
    let out = run_prints(r#"
        fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_bytes_" + System.nanoTime() + ".bin")
            file.writeBytes(byteArrayOf(1, 2, 3, 4))
            val bytes = file.readBytes()
            println(bytes.joinToString(","))
            file.delete()
        }
    "#);
    assert_eq!(out, &["1,2,3,4"]);
}

#[test]
fn test_kotlin_io_copy_to_duplicate_file() {
    let out = run_prints(r#"
        fun main() {
            val dir = java.io.File(java.lang.System.getProperty("java.io.tmpdir"))
            val src = java.io.File(dir, "vybe_io_copy_src_" + System.nanoTime() + ".txt")
            val dst = java.io.File(dir, "vybe_io_copy_dst_" + System.nanoTime() + ".txt")
            src.writeText("copy")
            val copied = src.copyTo(dst, overwrite = true)
            println(copied.readText())
            println(src.readText() == dst.readText())
            src.delete()
            dst.delete()
        }
    "#);
    assert_eq!(out, &["copy", "true"]);
}

#[test]
fn test_kotlin_io_stream_copy_between_files() {
    let out = run_prints(r#"
        fun main() {
            val dir = java.io.File(java.lang.System.getProperty("java.io.tmpdir"))
            val src = java.io.File(dir, "vybe_io_stream_src_" + System.nanoTime() + ".txt")
            val dst = java.io.File(dir, "vybe_io_stream_dst_" + System.nanoTime() + ".txt")
            src.writeText("stream")
            src.inputStream().use { input ->
                dst.outputStream().use { output ->
                    input.copyTo(output)
                }
            }
            println(dst.readText())
            src.delete()
            dst.delete()
        }
    "#);
    assert_eq!(out, &["stream"]);
}

#[test]
fn test_kotlin_io_rename_to_new_file() {
    let out = run_prints(r#"
        fun main() {
            val dir = java.io.File(java.lang.System.getProperty("java.io.tmpdir"))
            val src = java.io.File(dir, "vybe_io_rename_src_" + System.nanoTime() + ".txt")
            val dst = java.io.File(dir, "vybe_io_rename_dst_" + System.nanoTime() + ".txt")
            src.writeText("rename")
            val ok = src.renameTo(dst)
            println(ok)
            println(src.exists())
            println(dst.readText())
            dst.delete()
        }
    "#);
    assert_eq!(out, &["true", "false", "rename"]);
}

#[test]
fn test_kotlin_io_create_nested_directory_and_list_children() {
    let out = run_prints(r#"
        fun main() {
            val parent = java.io.File(java.lang.System.getProperty("java.io.tmpdir"), "vybe_io_dir_" + System.nanoTime())
            parent.mkdirs()
            val childA = java.io.File(parent, "a.txt")
            val childB = java.io.File(parent, "b.txt")
            childA.writeText("1")
            childB.writeText("2")
            val names = parent.listFiles()
            val joined = names.map { it.name }.sorted().joinToString(",")
            println(joined)
            println(parent.delete())
            childA.delete()
            childB.delete()
            parent.delete()
        }
    "#);
    assert_eq!(out, &["a.txt,b.txt", "false"]);
}

#[test]
fn test_kotlin_io_walk_top_down_includes_nested_files() {
    let out = run_prints(r#"
        fun main() {
            val parent = java.io.File(java.lang.System.getProperty("java.io.tmpdir"), "vybe_io_walk_" + System.nanoTime())
            val nested = java.io.File(parent, "nested")
            nested.mkdirs()
            java.io.File(parent, "root.txt").writeText("r")
            java.io.File(nested, "leaf.txt").writeText("l")
            val names = parent.walkTopDown().map { it.name }.toList().sorted()
            println(names.contains("nested"))
            println(names.contains("leaf.txt"))
            println(names.contains("root.txt"))
            java.io.File(parent, "root.txt").delete()
            java.io.File(nested, "leaf.txt").delete()
            nested.delete()
            parent.delete()
        }
    "#);
    assert_eq!(out, &["true", "true", "true"]);
}

#[test]
fn test_kotlin_io_walk_by_depth() {
    let out = run_prints(r#"
        fun main() {
            val parent = java.io.File(java.lang.System.getProperty("java.io.tmpdir"), "vybe_io_walk_depth_" + System.nanoTime())
            val level1 = java.io.File(parent, "level1")
            val level2 = java.io.File(level1, "level2")
            level2.mkdirs()
            java.io.File(level2, "leaf.txt").writeText("ok")
            val names = parent.walkBottomUp().map { it.name }.toList()
            println(names.contains("leaf.txt"))
            println(names.size)
            java.io.File(level2, "leaf.txt").delete()
            level2.delete()
            level1.delete()
            parent.delete()
        }
    "#);
    assert_eq!(out, &["true", "4"]);
}

#[test]
fn test_kotlin_io_path_properties() {
    let out = run_prints(r#"
        fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_props_" + System.nanoTime() + ".dat")
            file.writeText("p")
            println(file.name.endsWith(".dat"))
            println(file.extension)
            println(file.nameWithoutExtension.contains("vybe_io_props_"))
            file.delete()
        }
    "#);
    assert_eq!(out, &["true", "dat", "true"]);
}

#[test]
fn test_kotlin_io_absolute_and_parent_paths() {
    let out = run_prints(r#"
        fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_abs_" + System.nanoTime() + ".txt")
            file.writeText("x")
            val absolute = file.absolutePath
            val parent = file.parent
            println(absolute.startsWith(java.lang.System.getProperty("java.io.tmpdir")))
            println(parent != null)
            file.delete()
        }
    "#);
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_kotlin_io_file_delete() {
    let out = run_prints(r#"
        fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_delete_" + System.nanoTime() + ".txt")
            file.writeText("gone")
            val before = file.exists()
            val deleted = file.delete()
            println(before)
            println(deleted)
            println(file.exists())
        }
    "#);
    assert_eq!(out, &["true", "true", "false"]);
}

#[test]
fn test_kotlin_io_file_permissions_approx() {
    let out = run_prints(r#"
        fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_perm_" + System.nanoTime() + ".txt")
            file.writeText("perm")
            println(file.canRead())
            println(file.canWrite())
            file.delete()
        }
    "#);
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_kotlin_io_last_modified_is_time_like() {
    let out = run_prints(r#"
        fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_mtime_" + System.nanoTime() + ".txt")
            file.writeText("time")
            val mtime = file.lastModified()
            println(mtime > 0)
            file.delete()
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_kotlin_io_is_file_and_is_directory() {
    let out = run_prints(r#"
        fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_kind_" + System.nanoTime() + ".txt")
            file.writeText("kind")
            val dir = java.io.File(java.lang.System.getProperty("java.io.tmpdir"), "vybe_io_kind_dir_" + System.nanoTime())
            dir.mkdirs()
            println(file.isFile())
            println(dir.isDirectory())
            file.delete()
            dir.delete()
        }
    "#);
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_kotlin_io_empty_file_reads_empty_lines() {
    let out = run_prints(r#"
        fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_empty_" + System.nanoTime() + ".txt")
            file.writeText("")
            val bytes = file.readBytes()
            val lines = file.readLines()
            println(bytes.size)
            println(lines.size)
            file.delete()
        }
    "#);
    assert_eq!(out, &["0", "0"]);
}

#[test]
fn test_kotlin_io_file_exists_false_before_create() {
    let out = run_prints(r#"
        fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_missing_" + System.nanoTime() + ".txt")
            println(file.exists())
            file.writeText("now")
            println(file.exists())
            file.delete()
        }
    "#);
    assert_eq!(out, &["false", "true"]);
}

#[test]
fn test_kotlin_io_copy_to_fails_without_overwrite() {
    let out = run_prints(r#"
        fun main() {
            val dir = java.io.File(java.lang.System.getProperty("java.io.tmpdir"))
            val src = java.io.File(dir, "vybe_io_cfail_src_" + System.nanoTime() + ".txt")
            val dst = java.io.File(dir, "vybe_io_cfail_dst_" + System.nanoTime() + ".txt")
            src.writeText("src")
            dst.writeText("dst")
            try {
                src.copyTo(dst, overwrite = false)
                println("no_error")
            } catch (e: Exception) {
                println(e::class.simpleName)
            }
            src.delete()
            dst.delete()
        }
    "#);
    assert_eq!(out, &["FileAlreadyExistsException"]);
}

#[test]
fn test_kotlin_io_copy_to_can_force_overwrite() {
    let out = run_prints(r#"
        fun main() {
            val dir = java.io.File(java.lang.System.getProperty("java.io.tmpdir"))
            val src = java.io.File(dir, "vybe_io_cov_src_" + System.nanoTime() + ".txt")
            val dst = java.io.File(dir, "vybe_io_cov_dst_" + System.nanoTime() + ".txt")
            src.writeText("src")
            dst.writeText("dst")
            src.copyTo(dst, overwrite = true)
            println(dst.readText())
            src.delete()
            dst.delete()
        }
    "#);
    assert_eq!(out, &["src"]);
}

#[test]
fn test_kotlin_io_directory_file_count() {
    let out = run_prints(r#"
        fun main() {
            val parent = java.io.File(java.lang.System.getProperty("java.io.tmpdir"), "vybe_io_count_" + System.nanoTime())
            parent.mkdirs()
            java.io.File(parent, "a").mkdir()
            java.io.File(parent, "b.txt").writeText("b")
            java.io.File(parent, "c.txt").writeText("c")
            val files = parent.listFiles { f -> f.isFile }
            println(files.size)
            val dirs = parent.listFiles { f -> f.isDirectory }
            println(dirs.size)
            java.io.File(parent, "b.txt").delete()
            java.io.File(parent, "c.txt").delete()
            java.io.File(parent, "a").delete()
            parent.delete()
        }
    "#);
    assert_eq!(out, &["2", "1"]);
}

#[test]
fn test_kotlin_io_walk_with_depth_limit() {
    let out = run_prints(r#"
        fun main() {
            val parent = java.io.File(java.lang.System.getProperty("java.io.tmpdir"), "vybe_io_depth_" + System.nanoTime())
            val d1 = java.io.File(parent, "d1")
            val d2 = java.io.File(d1, "d2")
            d2.mkdirs()
            java.io.File(d2, "f1.txt").writeText("f")
            println(parent.walkTopDown().maxDepth(1).count())
            println(parent.walkTopDown().maxDepth(3).count())
            java.io.File(d2, "f1.txt").delete()
            d2.delete()
            d1.delete()
            parent.delete()
        }
    "#);
    assert_eq!(out, &["2", "4"]);
}

#[test]
fn test_kotlin_io_temp_file_names_are_distinct() {
    let out = run_prints(r#"
        fun main() {
            val a = java.io.File.createTempFile("vybe", "io")
            val b = java.io.File.createTempFile("vybe", "io")
            println(a.name != b.name)
            println(a.exists())
            println(b.exists())
            a.delete()
            b.delete()
        }
    "#);
    assert_eq!(out, &["true", "true", "true"]);
}

#[test]
fn test_kotlin_io_temporary_file_delete() {
    let out = run_prints(r#"
        fun main() {
            val temp = java.io.File.createTempFile("vybe_io_ttl", ".tmp")
            temp.writeText("ttl")
            temp.deleteOnExit()
            println(temp.exists())
            val deleted = temp.delete()
            println(deleted)
            println(temp.exists())
        }
    "#);
    assert_eq!(out, &["true", "true", "false"]);
}

#[test]
fn test_kotlin_io_reader_writer_round_trip() {
    let out = run_prints(r#"
        fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_rw_" + System.nanoTime() + ".txt")
            val writer = java.io.OutputStreamWriter(file.outputStream())
            writer.write("r")
            writer.write("w")
            writer.close()
            val reader = java.io.InputStreamReader(file.inputStream())
            val text = reader.readText()
            reader.close()
            println(text)
            file.delete()
        }
    "#);
    assert_eq!(out, &["rw"]);
}

#[test]
fn test_kotlin_io_file_uri_and_name() {
    let out = run_prints(r#"
        fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_uri_" + System.nanoTime() + ".txt")
            file.writeText("uri")
            val uri = file.toURI()
            println(uri.toString().endsWith(".txt"))
            println(file.name.startsWith("vybe_io_uri_"))
            file.delete()
        }
    "#);
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_kotlin_io_file_name_with_path_methods() {
    let out = run_prints(r#"
        fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_name_" + System.nanoTime() + ".txt")
            file.writeText("x")
            println(file.path.contains(file.name))
            println(file.absoluteFile.name)
            println(file.toPath().fileName.toString().endsWith(".txt"))
            file.delete()
        }
    "#);
    assert_eq!(out, &["true", "true", "true"]);
}

#[test]
fn test_kotlin_io_walk_sorted_names() {
    let out = run_prints(r#"
        fun main() {
            val parent = java.io.File(java.lang.System.getProperty("java.io.tmpdir"), "vybe_io_sorted_" + System.nanoTime())
            parent.mkdirs()
            java.io.File(parent, "c.txt").writeText("3")
            java.io.File(parent, "a.txt").writeText("1")
            java.io.File(parent, "b.txt").writeText("2")
            val names = parent.walk().filter { it.isFile }.map { it.name }.sorted().joinToString(",")
            println(names)
            java.io.File(parent, "a.txt").delete()
            java.io.File(parent, "b.txt").delete()
            java.io.File(parent, "c.txt").delete()
            parent.delete()
        }
    "#);
    assert_eq!(out, &["a.txt,b.txt,c.txt"]);
}

#[test]
fn test_kotlin_io_file_append_and_for_each_line() {
    let out = run_prints(r#"
        fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_append_lines_" + System.nanoTime() + ".txt")
            file.writeText("1\n")
            file.appendText("2\n")
            file.appendText("3\n")
            val total = file.readText().trim().split("\n").size
            val first = StringBuilder()
            file.forEachLine { first.append(it) }
            println(total)
            println(first.toString())
            file.delete()
        }
    "#);
    assert_eq!(out, &["3", "123"]);
}

#[test]
fn test_kotlin_io_file_parent_is_directory() {
    let out = run_prints(r#"
        fun main() {
            val file = java.io.File(java.lang.System.getProperty("java.io.tmpdir") + "/vybe_io_parent_" + System.nanoTime() + ".txt")
            file.writeText("x")
            val parent = file.parentFile
            println(parent.isDirectory())
            println(file.toPath().parent != null)
            file.delete()
        }
    "#);
    assert_eq!(out, &["true", "true"]);
}
