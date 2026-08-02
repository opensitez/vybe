// vybe-test: kotlin/java_io/test_java_io_data_output_input_roundtrip_int_long
// origin: languages/kotlin/tests/kotlin/test_java_io.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val sink = java.io.ByteArrayOutputStream()
            val writer = java.io.DataOutputStream(sink)
            writer.writeInt(12)
            writer.writeLong(99)
            writer.flush()
            val bytes = java.io.ByteArrayInputStream(sink.toByteArray())
            val reader = java.io.DataInputStream(bytes)
            __check((reader.readInt()).toString(), "12")
            __check((reader.readLong()).toString(), "99")
        }
