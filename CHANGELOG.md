# Changelog

All notable changes to FolderWatcher will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/),
and this project adheres to [Semantic Versioning](https://semver.org/).

## [1.0.0] - 2026-04-14

### Added

- Directory monitoring via `ReadDirectoryChangesW` with overlapped I/O
- On-demand COM callbacks to Access VBA via `GetObject` + `Application.Run`
- Automatic shutdown when Access closes (process handle monitoring)
- Bitness-aware executable selection (`#If Win64` in sample VBA module)
- Embed-and-extract pattern: store executables in `usys_Resources` table
- Custom application icon support via `SetAppIcon` and `AppIcon` database property
- Sample Access database with `StartWatching`, `StopWatching`, and `OnNewFile` callback
- EV code-signed executables (Grandjean & Braverman, Inc.)
