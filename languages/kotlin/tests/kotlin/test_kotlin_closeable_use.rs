use crate::helpers::run_prints;

#[test]
fn test_use_closes_closeable_resource_after_use() {
    let out = run_prints(
        r#"
        import java.io.Closeable

        class Tracker : Closeable {
            var closed = false
            override fun close() {
                closed = true
            }
        }

        fun main() {
            val tracker = Tracker()
            tracker.use {
                println(it.closed)
            }
            println(tracker.closed)
        }
    "#,
    );
    assert_eq!(out, &["false", "true"]);
}

#[test]
fn test_use_returns_lambda_result() {
    let out = run_prints(
        r#"
        import java.io.Closeable

        class Tracker : Closeable {
            var closed = false
            override fun close() {
                closed = true
            }
            fun value() = "done"
        }

        fun main() {
            val tracker = Tracker()
            val out = tracker.use { t ->
                t.value()
            }
            println(out)
            println(tracker.closed)
        }
    "#,
    );
    assert_eq!(out, &["done", "true"]);
}

#[test]
fn test_use_nests_and_closes_in_order() {
    let out = run_prints(
        r#"
        import java.io.Closeable

        val events = StringBuilder()

        class Tracker(val tag: String) : Closeable {
            override fun close() {
                events.append(tag)
            }
        }

        fun main() {
            Tracker("a").use {
                it
                Tracker("b").use {
                    it
                }
            }
            println(events.toString())
        }
    "#,
    );
    assert_eq!(out, &["ba"]);
}

#[test]
fn test_use_closes_on_exception() {
    let out = run_prints(
        r#"
        import java.io.Closeable

        class Tracker : Closeable {
            var closed = false
            override fun close() { closed = true }
        }

        fun main() {
            var closed = false
            val tracker = Tracker()
            try {
                tracker.use {
                    println("before")
                    throw IllegalStateException("x")
                }
            } catch (e: Exception) {
                println(e::class.simpleName)
                closed = tracker.closed
            }
            println(closed)
        }
    "#,
    );
    assert_eq!(out, &["before", "IllegalStateException", "true"]);
}

#[test]
fn test_byte_array_input_stream_use_block_reads() {
    let out = run_prints(
        r#"
        import java.io.ByteArrayInputStream

        fun main() {
            val stream = ByteArrayInputStream("abc".toByteArray())
            val text = stream.use { s ->
                val first = s.read()
                val second = s.read()
                s.available().toString() + "," + first.toChar() + "," + second.toChar()
            }
            println(text)
            try {
                println(stream.read())
            } catch (e: Exception) {
                println("closed")
            }
        }
    "#,
    );
    assert_eq!(out, &["1,b,c", "closed"]);
}

#[test]
fn test_file_writer_use_appends_and_closes() {
    let out = run_prints(
        r#"
        fun main() {
            val root = java.lang.System.getProperty("java.io.tmpdir")
            val file = java.io.File(root, "vybe_closeable_file_" + System.nanoTime() + ".txt")
            file.createNewFile()
            file.writeText("start")
            java.io.FileWriter(file, true).use { out ->
                out.write("-end")
            }
            val afterWrite = file.readText()
            val len = file.length().toString()
            file.delete()
            println(afterWrite)
            println(len == "8")
        }
    "#,
    );
    assert_eq!(out, &["start-end", "true"]);
}

#[test]
fn test_use_with_custom_resource_multiple_closes_prohibited() {
    let out = run_prints(
        r#"
        import java.io.Closeable

        class Counted : Closeable {
            var closeCount = 0
            override fun close() { closeCount++ }
        }

        fun main() {
            val tracked = Counted()
            tracked.use {
                println(tracked.closeCount)
            }
            println(tracked.closeCount)
            try {
                tracked.close()
                println("extra")
            } catch (e: Exception) {
                println("err")
            }
            println(tracked.closeCount)
        }
    "#,
    );
    assert_eq!(out, &["0", "1", "extra", "2"]);
}

#[test]
fn test_reader_reader_use() {
    let out = run_prints(
        r#"
        import java.io.BufferedReader
        import java.io.StringReader

        fun main() {
            val text = "alpha\nbeta"
            val reader = StringReader(text)
            val out = BufferedReader(reader).use { br ->
                br.readLine() + "|" + br.readLine()
            }
            println(out)
        }
    "#,
    );
    assert_eq!(out, &["alpha|beta"]);
}
