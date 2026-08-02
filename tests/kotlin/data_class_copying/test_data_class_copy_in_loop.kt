// vybe-test: kotlin/data_class_copying/test_data_class_copy_in_loop
// origin: languages/kotlin/tests/kotlin/test_data_class_copying.rs

data class Tally(val value: Int)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seed = listOf(1, 2, 3)
            val out = seed.fold(Tally(0)) { acc, next -> acc.copy(value = acc.value + next) }
            __check((out.value).toString(), "6")
        }
