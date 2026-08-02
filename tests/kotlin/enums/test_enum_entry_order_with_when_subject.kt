// vybe-test: kotlin/enums/test_enum_entry_order_with_when_subject
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Priority { FIRST, SECOND, THIRD }

        fun rank(priority: Priority): Int {
            return when (priority) {
                Priority.FIRST -> 1
                Priority.SECOND -> 2
                Priority.THIRD -> 3
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((rank(Priority.SECOND)).toString(), "2")
        }
