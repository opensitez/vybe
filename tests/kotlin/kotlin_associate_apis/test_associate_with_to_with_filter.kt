// vybe-test: kotlin/kotlin_associate_apis/test_associate_with_to_with_filter
// origin: languages/kotlin/tests/kotlin/test_kotlin_associate_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = mutableMapOf<String, Int>()
            listOf("x", "yy", "yyy", "yyyy").filter { it.length > 1 }.associateWithTo(out) { it.length }
            __check((out.size).toString(), "3")
            __check((out["yy"]).toString(), "2")
            __check((out.containsKey("x")).toString(), "false")
        }
