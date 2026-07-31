use crate::helpers::run_prints;

#[test]
fn test_path_file_name_and_parent() {
    let out = run_prints(r#"
        import java.nio.file.Paths

        fun main() {
            val path = Paths.get("/tmp", "alpha", "beta", "data.txt")
            println(path.fileName.toString())
            println(path.fileName.toString().length)
            println(path.parent?.fileName?.toString())
            println(path.root?.toString())
        }
    "#);
    assert_eq!(out, &["data.txt", "8", "beta", "/"]);
}

#[test]
fn test_files_temp_write_read_delete_roundtrip() {
    let out = run_prints(r#"
        import java.nio.file.Files
        import java.nio.file.Path

        fun main() {
            val tmp = Files.createTempFile("vybe", ".txt")
            Files.writeString(tmp, "hello")
            val text = Files.readString(tmp)
            println(text)
            val moved = Files.move(tmp, tmp.resolveSibling("vybe_moved_" + tmp.fileName.toString()), java.nio.file.StandardCopyOption.REPLACE_EXISTING)
            println(Files.exists(tmp))
            println(Files.exists(moved))
            println(Files.readString(moved))
            Files.delete(moved)
        }
    "#);
    assert_eq!(out, &["hello", "false", "true", "hello"]);
}

#[test]
fn test_path_resolve_and_normalize() {
    let out = run_prints(r#"
        import java.nio.file.Paths

        fun main() {
            val root = Paths.get("/tmp", "base")
            val child = root.resolve("a").resolve("../b.txt").normalize()
            println(child.toString())
            println(child.endsWith("b.txt"))
        }
    "#);
    assert_eq!(out, &["/tmp/base/b.txt", "true"]);
}

#[test]
fn test_files_copy_and_delete_if_exists() {
    let out = run_prints(r#"
        import java.nio.file.Files
        import java.nio.file.StandardCopyOption

        fun main() {
            val src = Files.createTempFile("vybe-copy-src", ".txt")
            Files.writeString(src, "alpha")
            val dst = Files.createTempFile("vybe-copy-dst", ".txt")
            Files.copy(src, dst, StandardCopyOption.REPLACE_EXISTING)
            val before = Files.readString(dst)
            val removed = Files.deleteIfExists(src)
            println(before)
            println(removed)
            println(Files.exists(src))
            Files.delete(dst)
        }
    "#);
    assert_eq!(out, &["alpha", "true", "false"]);
}

#[test]
fn test_files_size_and_exists_checks() {
    let out = run_prints(r#"
        import java.nio.file.Files

        fun main() {
            val path = Files.createTempFile("vybe-size", ".txt")
            Files.writeString(path, "kotlin")
            println(Files.exists(path))
            println(Files.isRegularFile(path))
            println(Files.size(path))
            Files.delete(path)
            println(Files.exists(path))
        }
    "#);
    assert_eq!(out, &["true", "true", "6", "false"]);
}

#[test]
fn test_directory_walk_with_filter() {
    let out = run_prints(r#"
        import java.nio.file.Files
        import java.nio.file.Paths

        fun main() {
            val base = Paths.get(java.lang.System.getProperty("java.io.tmpdir"), "vybe_walk_" + System.nanoTime().toString())
            val a = Files.createDirectories(base.resolve("a"))
            val b = base.resolve("a.txt")
            val c = base.resolve("b.log")
            Files.writeString(b, "one")
            Files.writeString(c, "two")
            val count = Files.list(base).filter { p -> Files.isRegularFile(p) }.count().toInt()
            println(count)
            val hasTxt = Files.newDirectoryStream(base, "*.txt").use {
                it.asSequence().count()
            }
            println(hasTxt)
            Files.delete(b)
            Files.delete(c)
            Files.delete(a)
            Files.delete(base)
        }
    "#);
    assert_eq!(out, &["2", "1"]);
}

#[test]
fn test_path_iterate_name_count() {
    let out = run_prints(r#"
        import java.nio.file.Paths

        fun main() {
            val path = Paths.get("/x/y/z/file.data")
            println(path.nameCount)
            var parts = ""
            for (part in path) {
                parts += part.fileName.toString() + "/"
            }
            println(parts)
            println(path.getName(1).toString())
        }
    "#);
    assert_eq!(out, &["3", "x/y/z/file.data/", "y"]);
}

#[test]
fn test_copy_to_multiple_targets_preserves_contents() {
    let out = run_prints(r#"
        import java.nio.file.Files
        import java.nio.file.StandardCopyOption

        fun main() {
            val src = Files.createTempFile("vybe-copy-a", ".txt")
            Files.writeString(src, "data")
            val d1 = Files.createTempFile("vybe-copy-b", ".txt")
            val d2 = Files.createTempFile("vybe-copy-c", ".txt")
            Files.copy(src, d1, StandardCopyOption.REPLACE_EXISTING)
            Files.copy(src, d2, StandardCopyOption.REPLACE_EXISTING)
            println(Files.readString(d1))
            println(Files.readString(d2))
            Files.delete(src)
            Files.delete(d1)
            Files.delete(d2)
        }
    "#);
    assert_eq!(out, &["data", "data"]);
}

#[test]
fn test_paths_to_file_and_delete_on_exit_flag() {
    let out = run_prints(r#"
        import java.nio.file.Files

        fun main() {
            val path = Files.createTempFile("vybe-exit", ".txt")
            path.toFile().deleteOnExit()
            println(path.toFile().exists())
            println(path.toFile().canWrite())
        }
    "#);
    assert_eq!(out, &["true", "true"]);
}
