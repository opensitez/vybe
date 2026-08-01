use crate::helpers::run_prints;

#[test]
fn test_java_io_byte_array_output_stream_starts_empty() {
    let out = run_prints(
        r#"
        fun main() {
            val out = java.io.ByteArrayOutputStream()
            println(out.size())
            println(out.toString())
        }
    "#,
    );
    assert_eq!(out, &["0", "",]);
}

#[test]
fn test_java_io_byte_array_output_stream_write_byte() {
    let out = run_prints(
        r#"
        fun main() {
            val stream = java.io.ByteArrayOutputStream()
            stream.write(65)
            println(stream.toString())
            println(stream.size())
        }
    "#,
    );
    assert_eq!(out, &["A", "1",]);
}

#[test]
fn test_java_io_byte_array_output_stream_write_byte_array() {
    let out = run_prints(
        r#"
        fun main() {
            val stream = java.io.ByteArrayOutputStream()
            val payload = "ok".toByteArray()
            stream.write(payload)
            println(stream.toString())
        }
    "#,
    );
    assert_eq!(out, &["ok"]);
}

#[test]
fn test_java_io_byte_array_output_stream_reset() {
    let out = run_prints(
        r#"
        fun main() {
            val stream = java.io.ByteArrayOutputStream()
            stream.write("before".toByteArray())
            stream.reset()
            stream.write("after".toByteArray())
            println(stream.toString())
        }
    "#,
    );
    assert_eq!(out, &["after"]);
}

#[test]
fn test_java_io_byte_array_output_stream_to_byte_array_len() {
    let out = run_prints(
        r#"
        fun main() {
            val stream = java.io.ByteArrayOutputStream()
            stream.write("abc".toByteArray())
            val data = stream.toByteArray()
            println(data.size)
            println(data[1])
        }
    "#,
    );
    assert_eq!(out, &["3", "98"]);
}

#[test]
fn test_java_io_byte_array_input_stream_read_single_bytes() {
    let out = run_prints(
        r#"
        fun main() {
            val input = java.io.ByteArrayInputStream("xyz".toByteArray())
            println(input.read())
            println(input.read())
            println(input.read())
            println(input.read())
        }
    "#,
    );
    assert_eq!(out, &["120", "121", "122", "-1"]);
}

#[test]
fn test_java_io_byte_array_input_stream_available_after_partial_read() {
    let out = run_prints(
        r#"
        fun main() {
            val input = java.io.ByteArrayInputStream("1234".toByteArray())
            println(input.read())
            println(input.available())
        }
    "#,
    );
    assert_eq!(out, &["49", "3"]);
}

#[test]
fn test_java_io_byte_array_input_stream_mark_and_reset_roundtrip() {
    let out = run_prints(
        r#"
        fun main() {
            val input = java.io.ByteArrayInputStream("abcd".toByteArray())
            println(input.markSupported())
            println(input.read())
            input.mark(3)
            println(input.read())
            println(input.read())
            input.reset()
            println(input.read())
        }
    "#,
    );
    assert_eq!(out, &["true", "97", "98", "99", "98"]);
}

#[test]
fn test_java_io_print_writer_writes_text() {
    let out = run_prints(
        r#"
        fun main() {
            val bytes = java.io.ByteArrayOutputStream()
            val printer = java.io.PrintWriter(bytes)
            printer.println("kotlin")
            printer.print("rocks")
            printer.flush()
            println(bytes.toString())
        }
    "#,
    );
    assert_eq!(out, &["kotlin\nrocks"]);
}

#[test]
fn test_java_io_print_writer_appendable_and_flush() {
    let out = run_prints(
        r#"
        fun main() {
            val bytes = java.io.ByteArrayOutputStream()
            val writer = java.io.PrintWriter(bytes)
            writer.append("first")
            writer.append('-')
            writer.println("second")
            writer.flush()
            println(bytes.toString())
        }
    "#,
    );
    assert_eq!(out, &["first-second\n"]);
}

#[test]
fn test_java_io_output_stream_writer_with_utf8() {
    let out = run_prints(
        r#"
        fun main() {
            val bytes = java.io.ByteArrayOutputStream()
            val writer = java.io.OutputStreamWriter(bytes, java.nio.charset.StandardCharsets.UTF_8)
            writer.write("ß")
            writer.flush()
            println(bytes.toString("UTF-8"))
        }
    "#,
    );
    assert_eq!(out, &["ß"]);
}

#[test]
fn test_java_io_input_stream_reader_line_reading() {
    let out = run_prints(
        r#"
        fun main() {
            val bytes = java.io.ByteArrayInputStream("a\nb\n".toByteArray())
            val reader = java.io.BufferedReader(java.io.InputStreamReader(bytes))
            println(reader.readLine())
            println(reader.readLine())
            println(reader.readLine() == null)
        }
    "#,
    );
    assert_eq!(out, &["a", "b", "true"]);
}

#[test]
fn test_java_io_string_reader_read_chars() {
    let out = run_prints(
        r#"
        fun main() {
            val reader = java.io.StringReader("k1")
            val buf = CharArray(2)
            println(reader.read(buf))
            println(buf.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["2", "k,1"]);
}

#[test]
fn test_java_io_string_writer_append_and_to_string() {
    let out = run_prints(
        r#"
        fun main() {
            val writer = java.io.StringWriter()
            writer.write("a")
            writer.append('b')
            writer.append("c")
            println(writer.toString())
        }
    "#,
    );
    assert_eq!(out, &["abc"]);
}

#[test]
fn test_java_io_char_array_writer_to_char_array() {
    let out = run_prints(
        r#"
        fun main() {
            val writer = java.io.CharArrayWriter()
            writer.write('x')
            writer.write("yz", 0, 2)
            val chars = writer.toCharArray()
            println(chars.size)
            println(String(chars))
        }
    "#,
    );
    assert_eq!(out, &["3", "xyz"]);
}

#[test]
fn test_java_io_buffered_writer_newline_and_to_string() {
    let out = run_prints(
        r#"
        fun main() {
            val stringWriter = java.io.StringWriter()
            val writer = java.io.BufferedWriter(stringWriter)
            writer.write("line1")
            writer.newLine()
            writer.write("line2")
            writer.flush()
            println(stringWriter.toString())
        }
    "#,
    );
    assert_eq!(out, &["line1\nline2"]);
}

#[test]
fn test_java_io_buffered_reader_mark_support() {
    let out = run_prints(
        r#"
        fun main() {
            val reader = java.io.BufferedReader(java.io.StringReader("12345"))
            println(reader.markSupported())
        }
    "#,
    );
    assert_eq!(out, &["true"]);
}

#[test]
fn test_java_io_buffered_reader_ready_before_read() {
    let out = run_prints(
        r#"
        fun main() {
            val reader = java.io.BufferedReader(java.io.StringReader("ok"))
            println(reader.ready())
            println(reader.read())
            println(reader.read())
            println(reader.read())
        }
    "#,
    );
    assert_eq!(out, &["true", "111", "107", "-1"]);
}

#[test]
fn test_java_io_pushback_input_stream_unread_and_read() {
    let out = run_prints(
        r#"
        fun main() {
            val base = java.io.ByteArrayInputStream("ab".toByteArray())
            val stream = java.io.PushbackInputStream(base)
            println(stream.read().toChar())
            stream.unread('c'.code)
            println(stream.read())
            println(stream.read())
        }
    "#,
    );
    assert_eq!(out, &["a", "99", "98"]);
}

#[test]
fn test_java_io_data_output_input_roundtrip_int_long() {
    let out = run_prints(
        r#"
        fun main() {
            val sink = java.io.ByteArrayOutputStream()
            val writer = java.io.DataOutputStream(sink)
            writer.writeInt(12)
            writer.writeLong(99)
            writer.flush()
            val bytes = java.io.ByteArrayInputStream(sink.toByteArray())
            val reader = java.io.DataInputStream(bytes)
            println(reader.readInt())
            println(reader.readLong())
        }
    "#,
    );
    assert_eq!(out, &["12", "99"]);
}

#[test]
fn test_java_io_data_output_input_boolean_utf() {
    let out = run_prints(
        r#"
        fun main() {
            val sink = java.io.ByteArrayOutputStream()
            val writer = java.io.DataOutputStream(sink)
            writer.writeBoolean(true)
            writer.writeUTF("hello")
            writer.flush()
            val reader = java.io.DataInputStream(java.io.ByteArrayInputStream(sink.toByteArray()))
            println(reader.readBoolean())
            println(reader.readUTF())
        }
    "#,
    );
    assert_eq!(out, &["true", "hello"]);
}

#[test]
fn test_java_io_print_stream_auto_flush() {
    let out = run_prints(
        r#"
        fun main() {
            val sink = java.io.ByteArrayOutputStream()
            val printer = java.io.PrintStream(sink, true)
            printer.println("one")
            println(sink.toString())
        }
    "#,
    );
    assert_eq!(out, &["one\n"]);
}

#[test]
fn test_java_io_buffered_input_stream_read_with_block() {
    let out = run_prints(
        r#"
        fun main() {
            val input = java.io.BufferedInputStream(java.io.ByteArrayInputStream("zz".toByteArray()))
            val buf = ByteArray(1)
            println(input.read(buf))
            println(buf[0].toInt())
        }
    "#,
    );
    assert_eq!(out, &["1", "122"]);
}

#[test]
fn test_java_io_filtered_input_stream_passthrough_count() {
    let out = run_prints(
        r#"
        fun main() {
            val input = java.io.BufferedInputStream(java.io.ByteArrayInputStream("aa".toByteArray()))
            val filtered = object : java.io.FilterInputStream(input) {}
            println(filtered.read())
            println(filtered.read())
            println(filtered.read())
        }
    "#,
    );
    assert_eq!(out, &["97", "97", "-1"]);
}

#[test]
fn test_java_io_line_number_reader_tracks_line() {
    let out = run_prints(
        r#"
        fun main() {
            val reader = java.io.LineNumberReader(java.io.StringReader("x\n y\n"))
            println(reader.readLine())
            println(reader.getLineNumber())
            println(reader.readLine())
            println(reader.getLineNumber())
        }
    "#,
    );
    assert_eq!(out, &["x", "1", " y", "2"]);
}

#[test]
fn test_java_io_input_stream_reader_resets_with_charset() {
    let out = run_prints(
        r#"
        fun main() {
            val bytes = java.io.ByteArrayInputStream("hi".toByteArray("UTF-8"))
            val reader = java.io.InputStreamReader(bytes, java.nio.charset.StandardCharsets.UTF_8)
            println(reader.ready())
            println(reader.read())
            println(reader.read())
            println(reader.read())
        }
    "#,
    );
    assert_eq!(out, &["true", "104", "105", "-1"]);
}

#[test]
fn test_java_io_output_stream_writer_flush_no_op_if_closed() {
    let out = run_prints(
        r#"
        fun main() {
            val sink = java.io.ByteArrayOutputStream()
            val writer = java.io.OutputStreamWriter(sink)
            writer.write("test")
            writer.flush()
            writer.close()
            println(sink.toString())
        }
    "#,
    );
    assert_eq!(out, &["test"]);
}

#[test]
fn test_java_io_string_writer_append_chain() {
    let out = run_prints(
        r#"
        fun main() {
            val writer = java.io.StringWriter()
            writer.append("a").append('b').append("c")
            println(writer.toString())
        }
    "#,
    );
    assert_eq!(out, &["abc"]);
}

#[test]
fn test_java_io_byte_array_input_stream_skip_count() {
    let out = run_prints(
        r#"
        fun main() {
            val input = java.io.ByteArrayInputStream("abcdef".toByteArray())
            println(input.skip(2))
            println(input.read())
        }
    "#,
    );
    assert_eq!(out, &["2", "99"]);
}

#[test]
fn test_java_io_sequence_input_stream_concatenates() {
    let out = run_prints(
        r#"
        fun main() {
            val first = java.io.ByteArrayInputStream("ab".toByteArray())
            val second = java.io.ByteArrayInputStream("cd".toByteArray())
            val seq = java.io.SequenceInputStream(first, second)
            println(seq.read().toChar())
            println(seq.read().toChar())
            println(seq.read().toChar())
            println(seq.read().toChar())
            println(seq.read())
        }
    "#,
    );
    assert_eq!(out, &["a", "b", "c", "d", "-1"]);
}

#[test]
fn test_java_io_print_stream_no_auto_flush_without_newline() {
    let out = run_prints(
        r#"
        fun main() {
            val sink = java.io.ByteArrayOutputStream()
            val printer = java.io.PrintStream(sink)
            printer.print("no_flush")
            printer.flush()
            println(sink.toString())
        }
    "#,
    );
    assert_eq!(out, &["no_flush"]);
}

#[test]
fn test_java_io_reader_read_text_entire_stream() {
    let out = run_prints(
        r#"
        fun main() {
            val text = "kotlin stream"
            val reader = java.io.StringReader(text)
            val writer = java.io.StringWriter()
            val buf = CharArray(4)
            while (true) {
                val count = reader.read(buf)
                if (count < 0) break
                writer.write(buf, 0, count)
            }
            println(writer.toString())
        }
    "#,
    );
    assert_eq!(out, &["kotlin stream"]);
}

#[test]
fn test_java_io_writer_error_state_is_false_after_success() {
    let out = run_prints(
        r#"
        fun main() {
            val sink = java.io.ByteArrayOutputStream()
            val writer = java.io.PrintWriter(sink)
            writer.print("x")
            writer.flush()
            println(writer.checkError())
        }
    "#,
    );
    assert_eq!(out, &["false"]);
}

#[test]
fn test_java_io_char_array_reader_skip_mark_reset() {
    let out = run_prints(
        r#"
        fun main() {
            val reader = java.io.CharArrayReader("abc".toCharArray())
            val one = reader.read()
            reader.mark(2)
            reader.skip(1)
            reader.reset()
            val afterReset = reader.read()
            println(one)
            println(afterReset)
        }
    "#,
    );
    assert_eq!(out, &["97", "98"]);
}

#[test]
fn test_java_io_filter_writer_passthrough() {
    let out = run_prints(
        r#"
        fun main() {
            val sink = java.io.StringWriter()
            val filter = object : java.io.FilterWriter(sink) {
                override fun write(i: Int) {
                    super.write(i)
                }
                override fun write(cbuf: CharArray, off: Int, len: Int) {
                    super.write(cbuf, off, len)
                }
                override fun write(str: String, off: Int, len: Int) {
                    super.write(str, off, len)
                }
                override fun flush() {
                    super.flush()
                }
                override fun close() {
                    super.close()
                }
            }
            filter.write("hello")
            filter.flush()
            println(sink.toString())
        }
    "#,
    );
    assert_eq!(out, &["hello"]);
}
