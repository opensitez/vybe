! vybe-test: fortran/keyword_arguments/kw_26
! origin: languages/fortran/tests/fortran/test_keyword_arguments.rs
module m
interface
subroutine s(x, y)
integer::x,y
end subroutine
end interface
end module m
program p
use m
call s(y=2, x=1)
end program p

subroutine s(x, y)
integer::x,y
end subroutine
