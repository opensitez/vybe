// vybe-test: kotlin/java_io/test_java_io_data_output_input_boolean_utf
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
            writer.writeBoolean(true)
            writer.writeUTF("hello")
            writer.flush()
            val reader = java.io.DataInputStream(java.io.ByteArrayInputStream(sink.toByteArray()))
            __check((reader.readBoolean()).toString(), "true")
            __check((reader.readUTF()).toString(), "hello")
        }
