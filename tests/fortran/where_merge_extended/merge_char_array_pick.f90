! vybe-test: fortran/where_merge_extended/merge_char_array_pick
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
character(len=1) :: a(3)=["A","B","C"]
character(len=1) :: b(3)=["X","Y","Z"]
logical :: m(3)=[.true.,.false.,.true.]
character(len=1) :: c(3)
c=merge(a,b,m)
if (trim(c(1)) /= "A") then
    print *, "FAIL: want [A] got [", c(1), "]"
    stop 1
end if
if (trim(c(2)) /= "Y") then
    print *, "FAIL: want [Y] got [", c(2), "]"
    stop 1
end if
if (trim(c(3)) /= "C") then
    print *, "FAIL: want [C] got [", c(3), "]"
    stop 1
end if
end program t
