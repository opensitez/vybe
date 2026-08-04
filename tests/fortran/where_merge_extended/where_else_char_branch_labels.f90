! vybe-test: fortran/where_merge_extended/where_else_char_branch_labels
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
integer :: s(3)=[1,5,12]
character(len=1) :: c(3)
where (s<3)
c="L"
elsewhere
c="H"
end where
if (trim(c(1)) /= "L") then
    print *, "FAIL: want [L] got [", c(1), "]"
    stop 1
end if
if (trim(c(2)) /= "H") then
    print *, "FAIL: want [H] got [", c(2), "]"
    stop 1
end if
if (trim(c(3)) /= "H") then
    print *, "FAIL: want [H] got [", c(3), "]"
    stop 1
end if
end program t
