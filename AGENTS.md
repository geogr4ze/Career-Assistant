# Development standards

- Preserve notification routing, quiet hours, limits, idempotency, approval gates,
  and production safeguards. Do not make live provider requests during tests.
- Document every named function, helper, method, and constructor using JSDoc.
  Cover parameters, requirements, return/postconditions, side effects and
  exceptions where applicable. Update documentation whenever behavior changes.
- Aim for 80–100 characters per line; avoid exceeding 120 except for unavoidable
  literals or external identifiers. Split long expressions and statement chains.
- Prefer descriptive names, clear control flow, consistent indentation, and
  cohesive functions. Comment the reason for decisions rather than obvious code.
- Keep credentials, recipient numbers, private IDs and local configuration out of
  Git. Do not log tokens or raw provider responses.
- Run `node tests/run.cjs` and inspect diffs before changing the live Apps Script.
  Report any verification limits; never deploy a formatting change automatically.
