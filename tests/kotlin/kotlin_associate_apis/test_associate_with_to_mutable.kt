// vybe-test: kotlin/kotlin_associate_apis/test_associate_with_to_mutable
// origin: languages/kotlin/tests/kotlin/test_kotlin_associate_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = mutableMapOf<String, Int>()
            listOf("a", "bb", "ccc").associateWithTo(out) { it.length }
            __check((out["bb"]).toString(), "2")
            __check((out.size).toString(), "3")
        }
