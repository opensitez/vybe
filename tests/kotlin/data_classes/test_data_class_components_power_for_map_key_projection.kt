// vybe-test: kotlin/data_classes/test_data_class_components_power_for_map_key_projection
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Tile(val row: Int, val col: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val tiles = listOf(Tile(0, 1), Tile(2, 3))
            val labels = tiles
                .map { it.component1() to it.component2() }
                .joinToString(";") { "${it.first},${it.second}" }
            __check((labels).toString(), "0,1;2,3")
        }
