use super::helpers::run_python;

#[test]
fn test_python_counter_common_cases() {
    let src = r#"
from collections import Counter
c = Counter('abbccc')
print(c['a'])
print(c['c'])
print(c.most_common(2))
"#;
    assert_eq!(run_python(src), vec!["1", "3", "[('c', 3), ('b', 2)]"]);
}

#[test]
fn test_python_counter_total_is_supported_fallback() {
    let src = r#"
import sys
from collections import Counter
c = Counter(a=2, b=3)
if hasattr(c, 'total'):
    print(c.total())
else:
    print(sum(c.values()) if sys.version_info >= (3, 0) else 0)
"#;
    assert_eq!(run_python(src), vec!["5"]);
}
