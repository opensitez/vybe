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
    procedure_results_01,
    "integer function f()
f=1
end function f
"
);
c!(
    procedure_results_02,
    "real function f()
f=1.0
end function f
"
);
c!(
    procedure_results_03,
    "complex function f()
f=(1.0,2.0)
end function f
"
);
c!(
    procedure_results_04,
    "logical function f()
f=.true.
end function f
"
);
c!(
    procedure_results_05,
    "character(len=3) function f()
f='abc'
end function f
"
);
c!(
    procedure_results_06,
    "integer function f(n)
integer::n
f=n
end function f
"
);
c!(
    procedure_results_07,
    "real function f(x)
real::x
f=x
end function f
"
);
c!(
    procedure_results_08,
    "type t
 integer::x
end type t
type(t) function f()
f%x=1
end function f
"
);
c!(
    procedure_results_09,
    "function f() result(r)
integer :: r
r=1
end function f
"
);
c!(
    procedure_results_10,
    "recursive integer function f() result(r)
r=1
end function f
"
);
