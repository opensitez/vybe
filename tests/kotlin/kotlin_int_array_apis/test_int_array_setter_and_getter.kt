// vybe-test: kotlin/kotlin_int_array_apis/test_int_array_setter_and_getter
// origin: languages/kotlin/tests/kotlin/test_kotlin_int_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = IntArray(3)
            a[0] = 4
            a[1] = 5
            a[2] = a[0] + a[1]
            __check((a[2]).toString(), "9")
        }
