// vybe-test: kotlin/generics/test_generic_out_variance_reader
// origin: languages/kotlin/tests/kotlin/test_generics.rs

interface Reader<out T> {
            fun read(): T
        }

        class NameReader : Reader<String> {
            override fun read(): String {
                return "ok"
            }
        }

        fun consume(reader: Reader<Any>): String {
            return reader.read().toString()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val reader: Reader<String> = NameReader()
            __check((consume(reader)).toString(), "ok")
        }
