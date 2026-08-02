// vybe-test: kotlin/kotlin_associate_apis/test_associate_with_to_sorted_result
// origin: languages/kotlin/tests/kotlin/test_kotlin_associate_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = mutableMapOf<Int, String>()
            listOf("x", "yy", "zzz").associateWithTo(out) { it.length.toString() }
            val keys = out.keys.toList().sorted()
            __check((keys.joinToString(",")).toString(), "1,2,3")
            __check((out[2]).toString(), "2")
        }
