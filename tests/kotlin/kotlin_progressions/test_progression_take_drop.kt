// vybe-test: kotlin/kotlin_progressions/test_progression_take_drop
// origin: languages/kotlin/tests/kotlin/test_kotlin_progressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = (1..10)
            val first = r.take(4)
            val remain = r.drop(4)
            __check((first.joinToString(",")).toString(), "[1, 2, 3, 4]")
            __check((remain.take(3).joinToString(",")).toString(), "[5, 6, 7]")
        }
