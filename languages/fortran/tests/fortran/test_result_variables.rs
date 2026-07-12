use super::helpers::compile_ok;
macro_rules! c {
    ($n:ident,$s:expr) => {
        #[test]
        fn $n() {
            compile_ok($s);
        }
    };
}
c!(
    result_variables_01,
    "integer function f() result(r)
r=1
end function f
"
);
c!(
    result_variables_02,
    "real function f() result(r)
r=1.0
end function f
"
);
c!(
    result_variables_03,
    "complex function f() result(r)
r=(1.0,2.0)
end function f
"
);
c!(
    result_variables_04,
    "character(len=3) function f() result(r)
r='abc'
end function f
"
);
c!(
    result_variables_05,
    "logical function f() result(r)
r=.true.
end function f
"
);
c!(
    result_variables_06,
    "type t
 integer::x
end type t
type(t) function f() result(r)
r%x=1
end function f
"
);
c!(
    result_variables_07,
    "integer function f(n) result(r)
integer::n
r=n
end function f
"
);
c!(
    result_variables_08,
    "real function f(x) result(r)
real::x
r=x
end function f
"
);
c!(
    result_variables_09,
    "recursive integer function f(n) result(r)
integer::n
r=1
end function f
"
);
c!(
    result_variables_10,
    "function f() result(r)
integer :: r
r=1
end function f
"
);
