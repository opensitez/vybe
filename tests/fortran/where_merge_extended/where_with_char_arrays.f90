! vybe-test: fortran/where_merge_extended/where_with_char_arrays
! origin: languages/fortran/tests/fortran/test_where_merge_extended.rs
program t
character(len=2) :: src(3)=["aa", "bb", "cc"]
character(len=2) :: dst(3)
where (src /= "bb")
dst = src
elsewhere
dst = "--"
end where
if (trim(trim(dst(1))) /= "aa") then
    print *, "FAIL: want [aa] got [", trim(dst(1)), "]"
    stop 1
end if
if (trim(trim(dst(2))) /= "--") then
    print *, "FAIL: want [--] got [", trim(dst(2)), "]"
    stop 1
end if
if (trim(trim(dst(3))) /= "cc") then
    print *, "FAIL: want [cc] got [", trim(dst(3)), "]"
    stop 1
end if
end program t
