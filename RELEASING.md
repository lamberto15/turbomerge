# Releasing

`turbomerge` is released with [Changesets](https://github.com/changesets/changesets) and the
[`release.yml`](./.github/workflows/release.yml) workflow. After the first version, nothing is
published from a developer machine.

## Day to day

For every user-visible change (feature, fix, behavior change), add a changeset to your PR:

```sh
pnpm changeset
```

Pick the bump (`patch` for fixes, `minor` for features, `major` for breaking changes; while the
version is `0.x`, breaking changes are `minor`), write a one-line summary for the changelog, and
commit the generated `.changeset/*.md`. Docs-, test- and CI-only PRs don't need one (label them
`no-changeset`).

## How a release happens

1. PRs with changesets are merged into `main`.
2. `release.yml` opens or updates a **"chore(release): version packages"** PR that bumps
   `package.json` and writes `CHANGELOG.md`. Nothing is published yet.
3. Merging that PR runs the publish job: `pnpm run verify` (the same checks as CI), then
   `changeset publish` publishes to npm through **trusted publishing** (OIDC: no npm token exists),
   with a provenance attestation, and creates the git tag and GitHub release.

Before merging the version PR, check that `CHANGELOG.md` reads well and that CI is green.

## One-time setup (maintainer)

First publish, by hand (npm only allows a trusted publisher on a package that already exists):

```sh
git switch main && git pull
pnpm install --frozen-lockfile
pnpm run verify
pnpm pack --dry-run        # check the file list: only what should ship
npm login
pnpm publish --access public
```

- [ ] npm account with **2FA** enabled.
- [ ] First publish done (commands above).
- [ ] npmjs.com → `turbomerge` → **Settings → Trusted publishing** → GitHub Actions:
  owner `lamberto15`, repository `turbomerge`, workflow `release.yml` (exact, case-sensitive),
  environment: leave empty. Do this right before the first CI release: npm doesn't
  validate the values, and a new trusted publisher expires if no publish succeeds soon after
  (see the npm docs for the current window).
- [ ] Same page: publishing access → **require 2FA and disallow tokens**.
- [ ] GitHub → Settings → Actions → General: workflow permissions **Read**, and tick
  **"Allow GitHub Actions to create and approve pull requests"**.
- [ ] Optional: repository secret `RELEASE_PAT` (fine-grained, this repo only, Contents and
  Pull requests: read and write) so the version PR runs CI.
- [ ] Branch ruleset on `main` requiring a PR and the `ci-ok` check; tag ruleset protecting
  release tags.
- [ ] The repository must be **public** for provenance.
- [ ] After the first release through CI, delete any old `NPM_TOKEN` secret.

## A bad release

`npm deprecate turbomerge@<version> "<reason>"`, then release a fix through the normal flow.
Unpublishing is only possible within 72 hours and breaks anyone who installed that version.
