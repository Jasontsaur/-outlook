"""Rule-based email classifier.

Rules live in rules.yaml (machine-readable) - see OUTLOOK_EMAIL_RULES.md for
the human-readable explanation of the same rules. Classification runs purely
on metadata already returned by outlook.search_emails (subject, sender), so
it costs zero extra Outlook calls; body_contains rules only apply when a
body string is explicitly passed in.
"""

import yaml


def load_rules(path):
    with open(path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    return [r for r in data.get("rules", []) if r.get("enabled", True)]


def _any_substring(haystack, keywords):
    haystack = (haystack or "").lower()
    return any(k.lower() in haystack for k in keywords)


def classify(email, rules, body=None):
    """email: dict with 'subject', 'sender_email', 'sender_name' (as returned
    by outlook.search_emails / list_recent_emails).

    Returns (destination, confidence, rule_id), or (None, 0.0, None) if no
    rule matched.
    """
    subject = email.get("subject", "")
    sender_email = email.get("sender_email") or ""
    sender_name = email.get("sender_name", "")

    for rule in rules:
        match = rule.get("match", {})

        if match.get("sender_domain_contains") and not _any_substring(
            sender_email, match["sender_domain_contains"]
        ):
            continue
        if match.get("subject_contains") and not _any_substring(subject, match["subject_contains"]):
            continue
        if match.get("sender_contains") and not (
            _any_substring(sender_email, match["sender_contains"])
            or _any_substring(sender_name, match["sender_contains"])
        ):
            continue
        if match.get("body_contains"):
            if body is None or not _any_substring(body, match["body_contains"]):
                continue

        return rule["destination"], rule.get("confidence", 0.5), rule["id"]

    return None, 0.0, None
