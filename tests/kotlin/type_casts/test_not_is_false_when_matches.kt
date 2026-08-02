// vybe-test: kotlin/type_casts/test_not_is_false_when_matches
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = "text"
            if (value !is Int) {
                __check(("not_int").toString(), "not_int")
            }
        }
