# Changelog

All notable changes to LatexToWord are documented here.

The project follows [Semantic Versioning](https://semver.org/).

## [Unreleased]

### Added

- Dependency-free Python package and command-line interface.
- Parsing for `$...$`, `$$...$$`, `\(...\)`, and `\[...\]` delimiters.
- Automated parser and command-line tests.
- GitHub Actions checks on Windows and Linux.
- GitHub issue forms, pull-request checklist, and Dependabot configuration.

### Changed

- The Word macro now converts only selected text and preserves unrelated formatting.
- Python and VBA discover paths at runtime rather than using a specific user's folders.
- `\tfrac` and `\dfrac` are normalized to `\frac` before Word conversion.
- Runtime errors now return a failure code and retain a diagnostic log.

### Removed

- Runtime logs and Microsoft Word lock files from version control.
- Third-party Python runtime dependencies.
- Document-wide font and paragraph formatting resets.

## [0.1.0] - 2024-10-17

- Initial development prototype combining Python and a Word VBA macro.
