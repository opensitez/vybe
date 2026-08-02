// vybe-test: kotlin/variance/test_variance_mutable_projection_read_write_split
// origin: languages/kotlin/tests/kotlin/test_variance.rs

fun copyValues(source: List<out Int>, target: MutableList<Int>) {
            source.forEach { target.add(it) }
            println(target.joinToString(","))
        }
        fun main() {
            val target = mutableListOf<Int>()
            copyValues(listOf(1, 2), target)
        }

