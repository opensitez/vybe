// vybe-test: kotlin/tuples/test_tuple_with_nullable_components_in_destructure
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pair: Pair<String?, String?> = Pair(null, "x")
            val (left, right) = pair
            __check((left == null).toString(), "true")
            __check((right.length).toString(), "1")
        }
