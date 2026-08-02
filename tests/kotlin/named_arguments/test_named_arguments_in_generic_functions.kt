// vybe-test: kotlin/named_arguments/test_named_arguments_in_generic_functions
// origin: languages/kotlin/tests/kotlin/test_named_arguments.rs

fun <T> pair(left: T, right: T): String = left.toString() + "," + right.toString()
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((pair<T=String>(left = "a", right = "b")).toString(), "a,b")
            __check((pair<Int>(left = 1, right = 2)).toString(), "1,2")
        }
