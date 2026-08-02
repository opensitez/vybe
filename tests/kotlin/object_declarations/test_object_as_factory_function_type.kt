// vybe-test: kotlin/object_declarations/test_object_as_factory_function_type
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

object Factory {
            fun build(prefix: String): String = prefix + "hash"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Factory.build("a")).toString(), "ahash")
            __check((Factory.build("b")).toString(), "bhash")
        }
