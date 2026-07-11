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
    misc_exec_seq_01,
    "program p
integer::x
x=1
x=x+1
print *,x
end program p
"
);
c!(
    misc_storage_assoc_02,
    "program p
integer::a(2),b
equivalence(a(1),b)
print *,1
end program p
"
);
c!(
    misc_equiv_03,
    "program p
integer::a,b
equivalence(a,b)
print *,1
end program p
"
);
c!(
    misc_common_04,
    "program p
integer::x
common /blk/ x
print *,x
end program p
"
);
c!(
    misc_cray_ptr_05,
    "program p
pointer (p, x)
integer x
print *,1
end program p
"
);
c!(
    misc_obsol_06,
    "program p
integer::n
assign 10 to n
go to n
10 continue
end program p
"
);
c!(
    misc_deleted_07,
    "program p
pause
end program p
"
);
c!(
    misc_proc_dep_08,
    "program p
integer::x
print *,x
end program p
"
);
c!(
    misc_conformance_09,
    "program p
implicit none
integer::x
x=1
print *,x
end program p
"
);
c!(
    misc_diag_10,
    "program p
integer::x
x=1
print *,x
end program p
"
);
c!(
    misc_constraint_11,
    "program p
integer::x
x=1
print *,x
end program p
"
);
c!(
    misc_syntax_12,
    "program p
print *,1
end program p
"
);
c!(
    misc_semantic_13,
    "program p
integer::x
x=1
print *,x
end program p
"
);
c!(
    misc_runtime_14,
    "program p
integer::x
x=1
print *,x
end program p
"
);
c!(
    misc_optbarrier_15,
    "program p
integer,volatile::x
x=1
print *,x
end program p
"
);
c!(
    misc_stop_16,
    "program p
stop
end program p
"
);
c!(
    misc_error_stop_17,
    "program p
error stop
end program p
"
);
c!(
    misc_return_18,
    "subroutine s()
return
end subroutine s
"
);
c!(
    misc_continue_19,
    "program p
continue
end program p
"
);
c!(
    misc_null_stmt_20,
    "program p
10 continue
end program p
"
);
