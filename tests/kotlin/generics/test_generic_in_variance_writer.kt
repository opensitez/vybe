// vybe-test: kotlin/generics/test_generic_in_variance_writer
// origin: languages/kotlin/tests/kotlin/test_generics.rs

interface Writer<in T> {
            fun write(value: T)
        }

        class Logger : Writer<Any> {
            var last: Any? = null

            override fun write(value: Any) {
                last = value
            }
        }

        fun emitInt(writer: Writer<Int>, value: Int): String {
            writer.write(value)
            return writer.toString()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val logger = Logger()
            val writer: Writer<Int> = logger
            writer.write(7)
            __check((logger.last).toString(), "7")
        }
