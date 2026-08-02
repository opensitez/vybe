// vybe-test: kotlin/object_expressions/test_object_expression_implements_interface_via_type_alias
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

interface Reader {
            fun read(): String
            fun fallback(): String = "none"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val reader: Reader = object : Reader {
                override fun read(): String = "ok"
            }
            __check((reader.read()).toString(), "ok")
            __check((reader.fallback()).toString(), "none")
        }
