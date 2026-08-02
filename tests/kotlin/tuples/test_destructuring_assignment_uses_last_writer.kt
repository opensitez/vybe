// vybe-test: kotlin/tuples/test_destructuring_assignment_uses_last_writer
// origin: languages/kotlin/tests/kotlin/test_tuples.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            var pair = Pair(1, 2)
            var (first, second) = pair
            first = 9
            pair = Pair(first, second + 1)
            __check((pair).toString(), "(9, 3)")
        }
