// vybe-test: kotlin/java_util_collections/test_java_collections_new_set_from_map
// origin: languages/kotlin/tests/kotlin/test_java_util_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = java.util.HashMap<String, Boolean>()
            val set = java.util.Collections.newSetFromMap(map)
            set.add("x")
            set.add("y")
            __check((set).toString(), "[x, y]")
            __check((map.size).toString(), "2")
        }
