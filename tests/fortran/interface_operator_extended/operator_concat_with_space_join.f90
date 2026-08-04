! vybe-test: fortran/interface_operator_extended/operator_concat_with_space_join
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gjoin
implicit none
type :: Token
character(len=6) :: word
end type Token
interface operator(//)
module procedure join_token
end interface
contains
function join_token(a, b) result(c)
type(Token), intent(in) :: a, b
type(Token) :: c
c%word = trim(a%word) // '-' // trim(b%word)
end function join_token
end module gjoin
program t
use gjoin
type(Token) :: x, y, z
x%word = 'ab'
y%word = 'cd'
z = x // y
if (trim(trim(z%word)) /= "ab-cd") then
    print *, "FAIL: want [ab-cd] got [", trim(z%word), "]"
    stop 1
end if
end program t
