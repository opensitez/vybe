use super::helpers::run_python;

#[test]
fn test_python_heapq_push_pop() {
    let src = r#"
import heapq
nums = [9, 1, 8, 3]
heapq.heapify(nums)
heapq.heappush(nums, 4)
first = heapq.heappop(nums)
second = heapq.heappushpop(nums, 2)
print(first)
print(second)
"#;
    assert_eq!(run_python(src), vec!["1", "2"]);
}

#[test]
fn test_python_heapq_nsmallest_nlargest() {
    let src = r#"
import heapq
nums = [7, 1, 6, 3, 2]
print(heapq.nsmallest(3, nums))
print(heapq.nlargest(2, nums))
"#;
    assert_eq!(run_python(src), vec!["[1, 2, 3]", "[7, 6]"]);
}
