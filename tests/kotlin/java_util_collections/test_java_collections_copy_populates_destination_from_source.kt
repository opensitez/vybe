// vybe-test: kotlin/java_util_collections/test_java_collections_copy_populates_destination_from_source
// origin: languages/kotlin/tests/kotlin/test_java_util_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source = java.util.ArrayList<Int>()
            source.add(1)
            source.add(2)
            source.add(3)
            val target = java.util.ArrayList<Int>(java.util.ArrayList<Int>(listOf(0, 0, 0, 0)))
            java.util.Collections.copy(target, source)
            __check((target).toString(), "[1, 2, 3, 0]")
        }
