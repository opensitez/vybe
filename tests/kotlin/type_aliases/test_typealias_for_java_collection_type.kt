// vybe-test: kotlin/type_aliases/test_typealias_for_java_collection_type
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias JavaMap = java.util.LinkedHashMap<String, Int>

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val counts: JavaMap = java.util.LinkedHashMap<String, Int>()
            counts["a"] = 1
            counts["b"] = 2
            counts.put("a", 3)
            __check((counts["a"]).toString(), "3")
            __check((counts.size).toString(), "2")
        }
