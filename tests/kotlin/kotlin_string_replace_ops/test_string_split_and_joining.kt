// vybe-test: kotlin/kotlin_string_replace_ops/test_string_split_and_joining
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_replace_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val parts = "a,b,c".split(",")
            __check((parts.size).toString(), "3")
            __check((parts[1]).toString(), "b")
        }
