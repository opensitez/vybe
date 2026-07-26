use super::helpers::run_python;

#[test]
fn test_python_functools_lru_cache() {
    let src = r#"
from functools import lru_cache

@lru_cache(maxsize=4)
def fib(n):
    if n < 2:
        return n
    return fib(n - 1) + fib(n - 2)

print(fib(10))
print(fib.cache_info().hits > 0)
"#;
    assert_eq!(run_python(src), vec!["55", "True"]);
}

#[test]
fn test_python_functools_singledispatch() {
    let src = r#"
from functools import singledispatch

@singledispatch
def show(x):
    return 'other'

@show.register(int)
def _(x):
    return 'int'

@show.register(str)
def _(x):
    return 'str'

print(show(1))
print(show('a'))
"#;
    assert_eq!(run_python(src), vec!["int", "str"]);
}
