// vybe-test: kotlin/named_arguments/test_named_arguments_named_arguments_internally_consistent
// origin: languages/kotlin/tests/kotlin/test_named_arguments.rs

fun total(one: Int = 1, two: Int = 2, three: Int = 3): Int {
            return one + two + three
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((total()).toString(), "6")
            __check((total(two = 10)).toString(), "12")
            __check((total(three = 7, one = 1, two = 2)).toString(), "10")
        }
