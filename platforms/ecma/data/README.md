# Vendored IANA time zone data

## `zone.tab`

Maps ISO 3166-1 alpha-2 country codes to IANA time zone identifiers. Backs
`Intl.Locale.prototype.getTimeZones` (ECMA-402 §14.3).

- **Upstream:** IANA Time Zone Database — <https://www.iana.org/time-zones>
  (release tarballs at <https://data.iana.org/time-zones/releases/>)
- **tzdb version:** 2026b
- **SHA-256 (unmodified file):** `4d8e389e5f4b0ec0466d5b14f42e5dfb0308c4376165fcf478339afd9ddcb00c`
- **License:** **PUBLIC DOMAIN.** The dedication is carried inside the file
  itself — *"This file is in the public domain, so clarified as of 2009-05-17
  by Arthur David Olson."* The tzdb distribution's `LICENSE` states the same
  for the whole data set: *"Unless specified below, all files in the tz code
  and data (including this LICENSE file) are in the public domain."* The only
  exceptions named there are `date.c`, `newstrftime.3` and `strftime.c`
  (BSD-3-Clause), which are C sources and are **not** vendored here.

No attribution is legally required. The file is stored unmodified so the
checksum above can be verified against the upstream release.

### Why this file is vendored at all

Neither of our timezone dependencies carries the region mapping:

- `chrono-tz` vendors the tzdb *rule* files (`africa`, `asia`, `europe`, …),
  which describe when offsets change. It has no `zone.tab` and no country API.
- ICU's region data (`icu_time::zone::windows`) maps *Windows* zone names plus
  a region to BCP-47 ids — the reverse direction, for a different purpose.

ECMA-402 references this file by name when specifying which identifiers are
primary, so it is the spec's own source, not a workaround.

### Version skew

`chrono-tz` currently vendors tzdb **2025b** while this table is **2026b**.
`timezone.rs::identifiers_for_region` validates every entry against
`chrono_tz::TZ_VARIANTS` and drops unknown zones, so the module can never
return an identifier its own `offset()` would fail to resolve. Re-sync this
file when `chrono-tz` bumps its tzdb release.

### Upgrading

```sh
curl -O https://data.iana.org/time-zones/releases/tzdata<VERSION>.tar.gz
tar xzf tzdata<VERSION>.tar.gz zone.tab
mv zone.tab platforms/ecma/data/zone.tab
# update the version + SHA-256 above
```

`zone.tab` is marked "deprecated" upstream in favour of `zone1970.tab`, whose
first column is a comma-separated country list for zones shared across
countries. `identifiers_for_region` parses both shapes, so either file can be
dropped in; `zone.tab` is used because ECMA-402 names it explicitly.
