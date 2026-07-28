"""Generate a salted PBKDF2-SHA256 hash for an app account.

Usage:
    python tools/hash_password.py

Paste the printed line into the [users] section of .streamlit/secrets.toml
(local) or into your host's secrets manager. The plaintext password is never
written to disk and must never be committed.
"""

import getpass
import hashlib
import os

ITERATIONS = 240_000


def main():
    username = input("Username: ").strip()
    password = getpass.getpass("Password: ")
    confirm = getpass.getpass("Confirm password: ")
    if password != confirm:
        raise SystemExit("Passwords do not match.")
    if len(password) < 10:
        print("Warning: short passwords are easy to brute force.")

    salt = os.urandom(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, ITERATIONS)
    print("\nAdd this under [users] in your secrets file:\n")
    print(f'{username} = "pbkdf2_sha256${ITERATIONS}${salt.hex()}${digest.hex()}"')


if __name__ == "__main__":
    main()
