! vybe-test: fortran/keyword_arguments/kw_09
! origin: languages/fortran/tests/fortran/test_keyword_arguments.rs
subroutine s(x)
integer::x
end
program p
call s(x=1)
end program p
