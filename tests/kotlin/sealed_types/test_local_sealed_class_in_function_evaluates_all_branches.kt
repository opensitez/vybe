// vybe-test: kotlin/sealed_types/test_local_sealed_class_in_function_evaluates_all_branches
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

fun classify(flag: Boolean): String {
            sealed class LocalResult {
                class Yes(val label: String) : LocalResult()
                class No : LocalResult()
            }

            val result: LocalResult = if (flag) LocalResult.Yes("ok") else LocalResult.No()
            return when (result) {
                is LocalResult.Yes -> result.label
                is LocalResult.No -> "no"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((classify(true)).toString(), "ok")
            __check((classify(false)).toString(), "no")
        }
