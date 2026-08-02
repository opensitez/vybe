// vybe-test: kotlin/object_declarations/test_object_can_be_used_as_factory_for_function_type
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

object Builder : (String, String) -> String {
            override fun invoke(left: String, right: String): String {
                return left + right
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: (String, String) -> String = Builder
            __check((value("a", "b")).toString(), "ab")
            __check((Builder.invoke("c", "d")).toString(), "cd")
        }
