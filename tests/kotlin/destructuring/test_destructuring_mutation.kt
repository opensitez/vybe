// vybe-test: kotlin/destructuring/test_destructuring_mutation
// origin: languages/kotlin/tests/kotlin/test_destructuring.rs

var a = 1
var b = 2
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val pair = Pair(a, b)
val (first, second) = pair
__check((first + second).toString(), "3")
}
