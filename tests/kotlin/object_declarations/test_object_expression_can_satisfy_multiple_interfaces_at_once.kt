// vybe-test: kotlin/object_declarations/test_object_expression_can_satisfy_multiple_interfaces_at_once
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

interface Named {
            fun name(): String
        }

        interface Valued {
            fun value(): Int
        }

        fun make(flag: String): Any {
            return object : Named, Valued {
                override fun name(): String = flag
                override fun value(): Int = 7
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item = make("ok")
            __check(((item as Named).name()).toString(), "ok")
            __check(((item as Valued).value()).toString(), "7")
        }
