! vybe-test: fortran/interface_operator_extended/operator_concat_three_tokens
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module g3tok
implicit none
type :: Tok
character(len=4) :: s
end type Tok
interface operator(//)
module procedure cat_tok
end interface
contains
function cat_tok(a, b) result(c)
type(Tok), intent(in) :: a, b
type(Tok) :: c
c%s = trim(a%s) // trim(b%s)
end function cat_tok
end module g3tok
program t
use g3tok
type(Tok) :: a, b, c, d
a%s = 'a'
b%s = 'b'
c%s = 'c'
d = (a // b) // c
if (trim(trim(d%s)) /= "abc") then
    print *, "FAIL: want [abc] got [", trim(d%s), "]"
    stop 1
end if
end program t
