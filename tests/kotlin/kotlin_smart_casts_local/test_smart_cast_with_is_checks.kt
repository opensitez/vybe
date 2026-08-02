// vybe-test: kotlin/kotlin_smart_casts_local/test_smart_cast_with_is_checks
// origin: languages/kotlin/tests/kotlin/test_kotlin_smart_casts_local.rs

fun describe(value: Any): String {
            return if (value is String) {
                "str:" + value.length
            } else if (value is Int) {
                "int:" + value
            } else {
                "other"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((describe("xy")).toString(), "str:2")
            __check((describe(7)).toString(), "int:7")
            __check((describe(true)).toString(), "other")
        }
