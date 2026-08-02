// vybe-test: kotlin/kotlin_char_apis/test_char_case_transform
// origin: languages/kotlin/tests/kotlin/test_kotlin_char_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = 'a'
            __check((c.uppercaseChar()).toString(), "A")
            __check((c.lowercaseChar()).toString(), "a")
            __check(('ß'.uppercase()).toString(), "SS")
            __check(('A'.uppercase().toString()).toString(), "A")
        }
