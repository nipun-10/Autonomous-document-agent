"""
Test script: demonstrates both standard and complex agent requests.
Run AFTER starting the API server: uvicorn agent:app --reload
"""
import httpx
import json
import asyncio
from pathlib import Path

BASE_URL = "http://localhost:8000"

# ── Test 1: Standard business request ─────────────────────────────────────────
STANDARD_REQUEST = (
    "Create a project plan for launching a new mobile banking app "
    "for a mid-sized regional bank. Include timeline, resource allocation, "
    "risk management, and budget overview."
)

# ── Test 2: Complex / ambiguous request ───────────────────────────────────────
COMPLEX_REQUEST = (
    "We need something for the thing we discussed last week about improving "
    "our team's output. There might be some process issues and also the new "
    "tool we're considering. Make it professional — the board might see it. "
    "Oh, and include the numbers somehow. We want it ready by tomorrow."
)


async def call_agent(label: str, request: str):
    print(f"\n{'='*60}")
    print(f"TEST: {label}")
    print(f"REQUEST: {request}")
    print("="*60)

    async with httpx.AsyncClient(timeout=120) as client:
        resp = await client.post(f"{BASE_URL}/agent", json={"request": request})
        resp.raise_for_status()
        data = resp.json()

    print(f"\n✅ REQUEST ID   : {data['request_id']}")
    print(f"📌 GOAL         : {data['interpreted_goal']}")
    print(f"📋 ASSUMPTIONS  :")
    for a in data["assumptions"]:
        print(f"   • {a}")
    print(f"\n📝 TASK LIST    :")
    for t in data["task_list"]:
        print(f"   [{t['id']}] {t['title']} — {t['status']}")
    print(f"\n📄 SUMMARY      :\n{data['execution_summary']}")
    print(f"\n📁 DOCUMENT     : {data['document_name']}")
    print(f"⏱️  COMPLETED    : {data['completed_at']}")

    # Download the file
    async with httpx.AsyncClient(timeout=30) as client:
        dl = await client.get(f"{BASE_URL}/download/{data['document_name']}")
        dl.raise_for_status()
        out = Path(f"downloaded_{data['document_name']}")
        out.write_bytes(dl.content)
        print(f"⬇️  DOWNLOADED   : {out} ({len(dl.content):,} bytes)")

    return data


async def main():
    print("\n🤖 AUTONOMOUS AI AGENT — TEST SUITE")
    print("Requires: uvicorn agent:app running on :8000")

    await call_agent("Standard Business Request", STANDARD_REQUEST)
    await call_agent("Complex / Ambiguous Request", COMPLEX_REQUEST)

    print("\n\n✅ All tests complete.")


if __name__ == "__main__":
    asyncio.run(main())
