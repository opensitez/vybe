use crate::helpers::run_prints;

#[test]
fn test_int_range_basic_boundaries() {
    let out = run_prints(r#"
        fun main() {
            val r = 1..5
            println(r.first)
            println(r.last)
            println(r.step)
            println(r.count())
        }
    "#);
    assert_eq!(out, &["1", "5", "1", "5"]);
}

#[test]
fn test_closed_range_contains() {
    let out = run_prints(r#"
        fun main() {
            val r = 3..7
            println(3 in r)
            println(8 in r)
            println(r.contains(5))
            println(r.contains(2))
        }
    "#);
    assert_eq!(out, &["true", "false", "true", "false"]);
}

#[test]
fn test_range_step_increments() {
    let out = run_prints(r#"
        fun main() {
            val r = (1..10 step 3)
            println(r.toList().joinToString(","))
            println(r.last)
        }
    "#);
    assert_eq!(out, &["1,4,7,10", "10"]);
}

#[test]
fn test_range_down_to_and_reversed() {
    let out = run_prints(r#"
        fun main() {
            val r = 10 downTo 4
            println(r.toList().joinToString(","))
            println(r.first)
            println(r.last)
            val asc = r.reversed()
            println(asc.toList().joinToString(","))
        }
    "#);
    assert_eq!(out, &["10,9,8,7,6,5,4", "10", "4", "4,5,6,7,8,9,10"]);
}

#[test]
fn test_empty_range_from_invalid_step() {
    let out = run_prints(r#"
        fun main() {
            val r = 5 downTo 10
            println(r.isEmpty())
            println(r.toList().isEmpty())
        }
    "#);
    assert_eq!(out, &["true", "true"]);
}

#[test]
fn test_char_range_projection() {
    let out = run_prints(r#"
        fun main() {
            val r = 'a'..'e'
            println(r.count())
            println(r.first())
            println(r.last())
            println(r.contains('c'))
            println(('d' in r).toString())
        }
    "#);
    assert_eq!(out, &["5", "a", "e", "true", "true"]);
}

#[test]
fn test_long_range_sum_projection() {
    let out = run_prints(r#"
        fun main() {
            val r = 2L..6L
            val sum = r.fold(0L) { acc, n -> acc + n }
            println(sum)
            println(r.any { it == 4L })
            println(r.all { it >= 2 })
        }
    "#);
    assert_eq!(out, &["20", "true", "true"]);
}

#[test]
fn test_open_ended_range_like_until() {
    let out = run_prints(r#"
        fun main() {
            val r = 1 until 4
            println(r.toList().joinToString(","))
            val step = (1 until 9 step 2)
            println(step.joinToString(","))
        }
    "#);
    assert_eq!(out, &["1,2,3", "1,3,5,7"]);
}

#[test]
fn test_range_map_projection() {
    let out = run_prints(r#"
        fun main() {
            val scaled = (1..4).map { it * 2 }
            val names = (1..3).map { "p" + it }
            println(scaled.joinToString(","))
            println(names.joinToString(","))
        }
    "#);
    assert_eq!(out, &["2,4,6,8", "p1,p2,p3"]);
}

#[test]
fn test_range_with() {
    let out = run_prints(r#"
        fun main() {
            val r = 1..10
            val firstHalf = r.take(3)
            val after = r.drop(7)
            println(firstHalf.joinToString(","))
            println(after.joinToString(","))
        }
    "#);
    assert_eq!(out, &["1,2,3", "8,9,10"]);
}

#[test]
fn test_when_with_range_subject() {
    let out = run_prints(r#"
        fun main() {
            val v = 4
            val label = when (v) {
                in 1..3 -> "small"
                in 4..6 -> "mid"
                else -> "big"
            }
            println(label)
            println(v in 1..10)
        }
    "#);
    assert_eq!(out, &["mid", "true"]);
}
