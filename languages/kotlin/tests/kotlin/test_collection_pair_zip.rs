use crate::helpers::run_prints;

#[test]
fn test_zip_aligns_minimum_length() {
    let out = run_prints(
        r#"
        fun main() {
            val left = listOf(1, 2, 3, 4)
            val right = listOf("a", "b")
            println(left.zip(right).joinToString("|") { "${it.first}${it.second}" })
            println(left.zip(right).size)
        }
    "#,
    );
    assert_eq!(out, &["1a|2b", "2"]);
}

#[test]
fn test_zip_with_transform() {
    let out = run_prints(
        r#"
        fun main() {
            val left = listOf(2, 4, 6)
            val right = listOf(1, 3, 5)
            val out = left.zip(right) { a, b -> a * b }
            println(out.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["2,12,30"]);
}

#[test]
fn test_zip_with_next_neighbor_pairs() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = listOf(1, 2, 3, 4)
            val pairs = nums.zipWithNext().joinToString("|") { "${it.first}:${it.second}" }
            println(pairs)
        }
    "#,
    );
    assert_eq!(out, &["1:2|2:3|3:4"]);
}

#[test]
fn test_zip_with_next_transform() {
    let out = run_prints(
        r#"
        fun main() {
            val nums = listOf(1, 2, 3, 4)
            val mapped = nums.zipWithNext { a, b -> a + b }
            println(mapped.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["3,5,7"]);
}

#[test]
fn test_unzip_restores_components() {
    let out = run_prints(
        r#"
        fun main() {
            val source = listOf(1 to "a", 2 to "b", 3 to "c")
            val (nums, chars) = source.unzip()
            println(nums.joinToString(","))
            println(chars.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1,2,3", "a,b,c"]);
}

#[test]
fn test_zip_and_unzip_round_trip() {
    let out = run_prints(
        r#"
        fun main() {
            val left = listOf(9, 8)
            val right = listOf("x", "y")
            val pair = left.zip(right)
            val back = pair.unzip()
            println(back.first.joinToString(","))
            println(back.second.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["9,8", "x,y"]);
}

#[test]
fn test_zip_empty_inputs() {
    let out = run_prints(
        r#"
        fun main() {
            println(listOf<Int>().zip(listOf("a")).isEmpty())
            println(listOf<Int>().zipWithNext().isEmpty())
        }
    "#,
    );
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_sequence_zip_with_extra_source() {
    let out = run_prints(
        r#"
        fun main() {
            val zipped = (1..5).asSequence().zip(listOf("a", "b", "c")) { n, s -> "$n$s" }.toList()
            println(zipped.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["1a,2b,3c"]);
}

#[test]
fn test_zip_preserves_laziness_on_sequence() {
    let out = run_prints(
        r#"
        fun main() {
            val counted = sequenceOf(1, 2).map { it + 1 }
            val zipped = counted.zip(sequenceOf(4, 5, 6)) { a, b -> a + b }
            println(zipped.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["6,8"]);
}

#[test]
fn test_product_via_zip() {
    let out = run_prints(
        r#"
        fun main() {
            val names = listOf("a", "b", "c")
            val out = names.zip(generateSequence(0) { it + 1 }) { name, i -> "$name$i" }
            println(out.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["a0,b1,c2"]);
}

#[test]
fn test_zip_with_shorter_right_iterable() {
    let out = run_prints(
        r#"
        fun main() {
            val words = listOf("one", "two", "three", "four")
            val nums = listOf(10)
            val zipped = words.zip(nums).joinToString(",") { "${it.first}:${it.second}" }
            println(zipped)
        }
    "#,
    );
    assert_eq!(out, &["one:10"]);
}

#[test]
fn test_flatten_and_zip_with_next_distinction() {
    let out = run_prints(
        r#"
        fun main() {
            val groups = listOf(listOf(1, 2), listOf(3, 4), listOf(5, 6))
            val zipped = groups.zipWithNext { a, b -> a.last() + b.first() }
            println(zipped.joinToString(","))
        }
    "#,
    );
    assert_eq!(out, &["5,9"]);
}
