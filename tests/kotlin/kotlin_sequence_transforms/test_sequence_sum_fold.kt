// vybe-test: kotlin/kotlin_sequence_transforms/test_sequence_sum_fold
// origin: languages/kotlin/tests/kotlin/test_kotlin_sequence_transforms.rs

fun main() {
            val seq = sequenceOf(1, 2, 3)
            var sum = 0
            for (v in seq) { sum += v }
            println(sum)
            val total = sequenceOf(10, 20, 30).fold(0) { acc, value -> acc + value }
            println(total)
        }

