# Contributing to ACLens

Thank you for your interest in contributing! This is an early alpha project and all kinds of contributions are welcome.

## Ways to Contribute

- 🐛 **Report bugs** — open an Issue with steps to reproduce
- 💡 **Suggest features** — open an Issue with the `enhancement` label
- 🔧 **Submit fixes** — open a Pull Request
- 📖 **Improve documentation** — README, comments, this guide

## Getting Started

1. Fork the repository
2. Clone your fork:
   ```powershell
   git clone https://github.com/YOUR-USERNAME/ACLens.git
   ```
3. Create a branch:
   ```powershell
   git checkout -b fix/my-bugfix
   # or
   git checkout -b feature/my-feature
   ```
4. Make your changes
5. Test on Windows 10 or 11 (and Server if available)
6. Commit with a clear message:
   ```powershell
   git commit -m "Fix: folder depth calculation for UNC paths"
   ```
7. Push and open a Pull Request against `main`

## Code Style

- PowerShell 5.1 compatible — no PS 7-only syntax
- Use `[void]$stringBuilder.Append(...)` for HTML generation
- All user-facing strings in English
- Keep GUI controls consistent with the existing DiskLens color palette
- Add a comment block for new functions

## Reporting Bugs

Please include:
- Windows version and PowerShell version (`$PSVersionTable`)
- Steps to reproduce
- Expected vs. actual behavior
- Error message if any (from the status bar or PowerShell console)

## Feature Requests

Please describe:
- The use case / problem you're solving
- Your proposed solution (if any)
- Any alternatives you considered

---

*By contributing, you agree that your contributions will be licensed under the MIT License.*
