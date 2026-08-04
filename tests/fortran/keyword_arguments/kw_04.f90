! vybe-test: fortran/keyword_arguments/kw_04
! origin: languages/fortran/tests/fortran/test_keyword_arguments.rs
subroutine s(x,y,z)
integer::x,y,z
end
program p
call s(1,y=2,z=3)
end program p
