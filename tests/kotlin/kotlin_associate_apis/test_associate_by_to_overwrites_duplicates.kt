// vybe-test: kotlin/kotlin_associate_apis/test_associate_by_to_overwrites_duplicates
// origin: languages/kotlin/tests/kotlin/test_kotlin_associate_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = linkedHashMapOf<Int, String>()
            listOf("ab", "ac", "b").associateByTo(out, { it.length }, { it })
            __check((out.size).toString(), "2")
            __check((out[2]).toString(), "b")
        }
