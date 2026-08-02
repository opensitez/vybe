// vybe-test: kotlin/kotlin_nothing_type/test_nothing_type_is_used_in_never_returning_flow
// origin: languages/kotlin/tests/kotlin/test_kotlin_nothing_type.rs

fun failNow(): Nothing = throw Exception("x")

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((try {
                failNow()
                "bad"
            } catch (e: Exception) {
                "caught"
            }).toString(), "caught")
        }
