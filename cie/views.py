import os
import tempfile
from django.shortcuts import render

# Adjust to your app structure, e.g.: from myapp.services.processor import process_zip
from .services.processor import process_zip


def index(request):
    context = {}

    if request.method == "POST" and request.FILES.get("file"):
        uploaded  = request.FILES["file"]
        threshold = int(request.POST.get("threshold", 50))

        with tempfile.NamedTemporaryFile(delete=False, suffix=".zip") as tmp:
            for chunk in uploaded.chunks():
                tmp.write(chunk)
            tmp_path = tmp.name

        try:
            result = process_zip(tmp_path, threshold_pct=threshold)
        except Exception as e:
            result = {
                "subject_counts": {},
                "subject_totals": {},
                "rows": [],
                "total_slow": 0,
                "processed": 0,
                "errors": [str(e)],
            }
        finally:
            os.unlink(tmp_path)

        counts = result["subject_counts"]
        totals = result["subject_totals"]

        # Pre-compute bar %  so template needs zero custom filters
        breakdown = [
            {
                "subject": s,
                "slow":    counts[s],
                "total":   totals.get(s, 0),
                "pct":     round(counts[s] / totals[s] * 100) if totals.get(s) else 0,
            }
            for s in counts
        ]

        context = {
            "subject_breakdown": breakdown,
            "rows":              result["rows"],
            "total_slow":        result["total_slow"],
            "processed":         result["processed"],
            "errors":            result["errors"],
            "threshold":         threshold,
        }

    return render(request, "index.html", context)