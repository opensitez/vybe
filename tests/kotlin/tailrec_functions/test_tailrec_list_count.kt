// vybe-test: kotlin/tailrec_functions/test_tailrec_list_count
// origin: languages/kotlin/tests/kotlin/test_tailrec_functions.rs

tailrec fun itemCount(items: List<Int>, idx: Int = 0, acc: Int = 0): Int {
            return if (idx >= items.size) acc else itemCount(items, idx + 1, acc + 1)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((itemCount(listOf(1, 2, 3, 4))).toString(), "4")
        }
