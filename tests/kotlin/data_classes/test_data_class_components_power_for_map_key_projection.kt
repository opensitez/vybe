// vybe-test: kotlin/data_classes/test_data_class_components_power_for_map_key_projection
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Tile(val row: Int, val col: Int)

        var __buf: String = ""

fun __p(s: String) {
    __buf = __buf + s + "\n"
}

fun __pr(s: String) {
    __buf = __buf + s
}

// The final `println` contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted. Written as two equality
// tests rather than trimming: `String.endsWith` is not implemented in Vybe's
// Kotlin (measured — `"ab\n".endsWith("\n")` throws "undefined is not
// callable"), and a harness that cannot run asserts nothing at all. The cargo
// helper split on "\n" and popped trailing empties, so the two forms were
// equivalent there too.
fun __check(want: String) {
    if (__buf != want && __buf != want + "\n") {
        println("FAIL: want [" + want + "] got [" + __buf + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val tiles = listOf(Tile(0, 1), Tile(2, 3))
            val labels = tiles
                .map { it.component1() to it.component2() }
                .joinToString(";") { "${it.first},${it.second}" }
            __p((labels).toString())
        
__check("0,1;2,3")
}
