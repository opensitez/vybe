! vybe-test: fortran/keyword_arguments/kw_15
! origin: languages/fortran/tests/fortran/test_keyword_arguments.rs
subroutine s(a,b)
integer::a,b
end
program p
call s(b=2, a=1)
end program p
