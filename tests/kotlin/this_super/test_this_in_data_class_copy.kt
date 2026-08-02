// vybe-test: kotlin/this_super/test_this_in_data_class_copy
// origin: languages/kotlin/tests/kotlin/test_this_super.rs

data class X(val v: Int)
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
    val x = X(1)
    val y = x.copy(v = x.v + 1)
    __check((y.v).toString(), "1")
}
