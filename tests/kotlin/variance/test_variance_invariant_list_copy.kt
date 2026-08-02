// vybe-test: kotlin/variance/test_variance_invariant_list_copy
// origin: languages/kotlin/tests/kotlin/test_variance.rs

fun copyAll(src: List<out Number>, dst: MutableList<in Number>) {
            src.forEach { dst.add(it.toInt()) }
            println(dst)
        }
        fun main() {
            val output = mutableListOf<Number>()
            copyAll(listOf(1, 2, 3), output)
        }

