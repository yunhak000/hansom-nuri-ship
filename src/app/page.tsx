"use client";

import { useEffect, useMemo, useState } from "react";
import { readExcelFile } from "@/lib/excel/read";
import { buildAggregateRows, downloadAggregateExcel, type TAggregateRow } from "@/lib/excel/writeAggregate";
import { buildCjGroupedRows, downloadCjUploadsZip } from "@/lib/excel/writeCJUploads";
import { readCjReplyFiles } from "@/lib/excel/readCJReply";
import { applyTracking, downloadOriginalWithTracking, downloadUnmatchedExcel } from "@/lib/excel/applyTrackingToOriginal";
import { clearJob, loadJob, saveJob, type TJobState } from "@/lib/db";
import { fingerprintFile, isSameFingerprint, type TFileFingerprint } from "@/lib/utils/hash";

type TRow = Record<string, any>;

type TStep = 1 | 2 | 3 | 4;

export default function HomePage() {
  const [step, setStep] = useState<TStep>(1);
  const [loading, setLoading] = useState<{ on: boolean; text: string }>({ on: false, text: "" });
  const [error, setError] = useState<string>("");

  // persisted job
  const [job, setJob] = useState<TJobState | null>(null);

  // derived
  const aggregateRows: TAggregateRow[] = useMemo(() => {
    if (!job?.originalRows) return [];
    return buildAggregateRows(job.originalRows);
  }, [job?.originalRows]);

  useEffect(() => {
    (async () => {
      const saved = await loadJob();
      if (saved) setJob(saved);
    })();
  }, []);

  const setBusy = (on: boolean, text = "") => setLoading({ on, text });

  const canStep2 = !!job?.originalRows?.length;
  const canStep3 = canStep2; // 원본 있어야 회신도 가능
  const canStep4 = canStep3; // 회신 업로드 후 활성화는 아래에서 실제로 제어

  const onReset = async () => {
    if (loading.on) return;
    await clearJob();
    setJob(null);
    setStep(1);
    setError("");
  };

  const onUploadOriginal = async (file: File | null) => {
    if (!file) return;
    try {
      setError("");
      setBusy(true, "원본 엑셀 읽는 중...");
      const { headers, rows } = await readExcelFile(file);

      const next: TJobState = {
        createdAt: new Date().toISOString(),
        originalFileName: file.name,
        originalHeaders: headers,
        originalRows: rows,
        uploadedReplyFiles: [],
      };

      await saveJob(next);
      setJob(next);
      setStep(2);
    } catch (e: any) {
      setError(e?.message ?? "원본 엑셀 처리 중 오류가 발생했습니다.");
    } finally {
      setBusy(false);
    }
  };

  const onDownloadAggregate = async () => {
    try {
      setError("");
      setBusy(true, "품목별 집계 엑셀 생성 중...");
      await downloadAggregateExcel(aggregateRows);
      setStep(2);
    } catch (e: any) {
      setError(e?.message ?? "집계 엑셀 생성 중 오류가 발생했습니다.");
    } finally {
      setBusy(false);
    }
  };

  const onDownloadCjZip = async () => {
    if (!job) return;
    try {
      setError("");
      setBusy(true, "CJ 업로드용 품목별 파일 생성(Zip) 중...");
      const groups = buildCjGroupedRows(job.originalHeaders, job.originalRows);
      await downloadCjUploadsZip(groups);
      setStep(3);
    } catch (e: any) {
      setError(e?.message ?? "CJ 업로드용 파일 생성 중 오류가 발생했습니다.");
    } finally {
      setBusy(false);
    }
  };

  const onUploadReplies = async (files: FileList | null) => {
    if (!job) return;
    if (!files || files.length === 0) return;

    try {
      setError("");

      // 중복 파일 필터링 + 경고
      const existing = job.uploadedReplyFiles ?? [];
      const newFingerprints: TFileFingerprint[] = [];
      const accepted: File[] = [];
      const dupNames: string[] = [];

      Array.from(files).forEach((f) => {
        const fp = fingerprintFile(f);
        const isDup = existing.some((x) => isSameFingerprint(x, fp)) || newFingerprints.some((x) => isSameFingerprint(x, fp));
        if (isDup) {
          dupNames.push(f.name);
          return;
        }
        newFingerprints.push(fp);
        accepted.push(f);
      });

      if (accepted.length === 0) {
        setError(`이미 업로드한 회신 파일입니다: ${dupNames.join(", ")}`);
        return;
      }
      if (dupNames.length > 0) {
        // 경고는 error 말고 안내로 하고 싶으면 toast로 바꾸면 됨 (지금은 간단히 error에 표시)
        setError(`일부 회신 파일은 이미 업로드되어 제외했습니다: ${dupNames.join(", ")}`);
      }

      setBusy(true, "CJ 회신 파일 읽는 중...");
      const { map } = await readCjReplyFiles(accepted);

      setBusy(true, "운송장번호 매핑 중...");
      const { updatedRows, unmatched, duplicates } = applyTracking(job.originalHeaders, job.originalRows, map);

      // job 저장 업데이트
      const next: TJobState = {
        ...job,
        originalRows: updatedRows,
        uploadedReplyFiles: [...existing, ...newFingerprints],
      };

      await saveJob(next);
      setJob(next);

      // step4로 이동 + 결과를 화면에 보여주기 위해 상태로 저장
      setLocalResult({
        unmatched,
        duplicates,
        matchedCount: map.size - unmatched.length,
        totalReplyCount: map.size,
      });
      setStep(4);
    } catch (e: any) {
      setError(e?.message ?? "회신 처리 중 오류가 발생했습니다.");
    } finally {
      setBusy(false);
    }
  };

  const [localResult, setLocalResult] = useState<{
    unmatched: Array<{ customerOrderNo: string; tracking: string }>;
    duplicates: Array<{ key: string; count: number }>;
    matchedCount: number;
    totalReplyCount: number;
  } | null>(null);

  const onDownloadFinal = async () => {
    if (!job) return;
    try {
      setError("");
      setBusy(true, "최종 원본 엑셀 생성 중...");
      await downloadOriginalWithTracking(job.originalHeaders, job.originalRows);
    } catch (e: any) {
      setError(e?.message ?? "최종 엑셀 생성 중 오류가 발생했습니다.");
    } finally {
      setBusy(false);
    }
  };

  const onDownloadUnmatched = async () => {
    if (!localResult) return;
    try {
      setError("");
      setBusy(true, "미매칭 목록 엑셀 생성 중...");
      await downloadUnmatchedExcel(localResult.unmatched);
    } catch (e: any) {
      setError(e?.message ?? "미매칭 엑셀 생성 중 오류가 발생했습니다.");
    } finally {
      setBusy(false);
    }
  };

  return (
    <main className="mx-auto max-w-4xl p-6 space-y-6">
      <header className="space-y-2">
        <h1 className="text-2xl font-bold">한섬누리 출고 엑셀 도구</h1>
        <p className="text-sm text-gray-600">원본 업로드 → 품목 집계/ CJ 업로드 파일 생성 → 회신 업로드 → 운송장 반영</p>
      </header>

      {/* Stepper */}
      <div className="flex gap-2 text-sm">
        {[
          { n: 1, label: "원본 업로드" },
          { n: 2, label: "산출물 생성" },
          { n: 3, label: "회신 업로드" },
          { n: 4, label: "최종 다운로드" },
        ].map((s) => (
          <div key={s.n} className={`flex-1 rounded-lg border p-3 text-center ${step === s.n ? "border-black font-semibold" : "border-gray-200"}`}>
            {s.n}. {s.label}
          </div>
        ))}
      </div>

      {/* Error */}
      {error && <div className="rounded-lg border border-red-300 bg-red-50 p-4 text-sm text-red-700">{error}</div>}

      {/* Loading Overlay */}
      {loading.on && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/30">
          <div className="rounded-2xl bg-white p-6 shadow-lg flex items-center gap-3">
            <div className="h-5 w-5 animate-spin rounded-full border-2 border-gray-300 border-t-black" />
            <div className="text-sm font-medium">{loading.text}</div>
          </div>
        </div>
      )}

      {/* Step 1 */}
      <section className="rounded-2xl border p-5 space-y-3">
        <h2 className="text-lg font-semibold">1) 원본 엑셀 업로드</h2>
        <input type="file" accept=".xlsx" disabled={loading.on} onChange={(e) => onUploadOriginal(e.target.files?.[0] ?? null)} />
        {job?.originalFileName && (
          <div className="text-sm text-gray-700">
            현재 작업 원본: <span className="font-medium">{job.originalFileName}</span>
          </div>
        )}

        <button
          className="rounded-lg bg-gray-900 px-4 py-2 text-sm text-white disabled:opacity-40"
          disabled={!job || loading.on}
          onClick={() => {
            if (confirm("모든 작업을 초기화하시겠습니까?")) {
              onReset();
            }
          }}
        >
          전체 리셋
        </button>
      </section>

      {/* Step 2 */}
      <section className="rounded-2xl border p-5 space-y-4">
        <h2 className="text-lg font-semibold">2) 산출물 생성</h2>

        <div className="flex flex-wrap gap-2">
          <button
            className="rounded-lg bg-black px-4 py-2 text-sm text-white disabled:opacity-40"
            disabled={!canStep2 || loading.on}
            onClick={onDownloadAggregate}
            title={!canStep2 ? "원본을 먼저 업로드하세요" : ""}
          >
            품목별 집계 엑셀 다운로드
          </button>

          <button
            className="rounded-lg bg-black px-4 py-2 text-sm text-white disabled:opacity-40"
            disabled={!canStep2 || loading.on}
            onClick={onDownloadCjZip}
            title={!canStep2 ? "원본을 먼저 업로드하세요" : ""}
          >
            CJ 업로드용 품목별 ZIP 다운로드
          </button>
        </div>

        <div className="text-sm text-gray-600">
          집계 건수: <span className="font-medium">{aggregateRows.length}</span>
        </div>
      </section>

      {/* Step 3 */}
      <section className="rounded-2xl border p-5 space-y-3">
        <h2 className="text-lg font-semibold">3) CJ 회신 엑셀 업로드(다중)</h2>
        <input type="file" accept=".xlsx" multiple disabled={!canStep3 || loading.on} onChange={(e) => onUploadReplies(e.target.files)} title={!canStep3 ? "원본을 먼저 업로드하세요" : ""} />
        <p className="text-sm text-gray-600">같은 파일을 다시 올리면 경고 후 제외됩니다.</p>
      </section>

      {/* Step 4 */}
      <section className="rounded-2xl border p-5 space-y-4">
        <h2 className="text-lg font-semibold">4) 결과 확인 & 다운로드</h2>

        {localResult ? (
          <div className="grid grid-cols-2 gap-3 text-sm">
            <div className="rounded-lg border p-3">
              회신 키 수: <span className="font-semibold">{localResult.totalReplyCount}</span>
            </div>
            <div className="rounded-lg border p-3">
              매핑 성공(추정): <span className="font-semibold">{localResult.matchedCount}</span>
            </div>
            <div className="rounded-lg border p-3">
              미매칭: <span className="font-semibold">{localResult.unmatched.length}</span>
            </div>
            <div className="rounded-lg border p-3">
              원본 중복키: <span className="font-semibold">{localResult.duplicates.length}</span>
            </div>
          </div>
        ) : (
          <div className="text-sm text-gray-600">회신 업로드 후 결과가 표시됩니다.</div>
        )}

        <div className="flex flex-wrap gap-2">
          <button className="rounded-lg bg-black px-4 py-2 text-sm text-white disabled:opacity-40" disabled={!job || loading.on} onClick={onDownloadFinal}>
            최종 원본 다운로드
          </button>

          <button
            className="rounded-lg bg-gray-800 px-4 py-2 text-sm text-white disabled:opacity-40"
            disabled={!localResult || localResult.unmatched.length === 0 || loading.on}
            onClick={onDownloadUnmatched}
            title={!localResult || localResult.unmatched.length === 0 ? "미매칭이 없습니다" : ""}
          >
            미매칭 목록 다운로드
          </button>
        </div>

        {localResult?.unmatched?.length ? (
          <div className="space-y-2">
            <div className="text-sm font-semibold">미매칭 리스트(상위 20개)</div>
            <div className="max-h-56 overflow-auto rounded-lg border">
              <table className="w-full text-sm">
                <thead className="sticky top-0 bg-white">
                  <tr>
                    <th className="border-b p-2 text-left">고객주문번호</th>
                    <th className="border-b p-2 text-left">운송장번호</th>
                  </tr>
                </thead>
                <tbody>
                  {localResult.unmatched.slice(0, 20).map((u) => (
                    <tr key={`${u.customerOrderNo}-${u.tracking}`}>
                      <td className="border-b p-2">{u.customerOrderNo}</td>
                      <td className="border-b p-2">{u.tracking}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        ) : null}

        {localResult?.duplicates?.length ? (
          <div className="space-y-2">
            <div className="text-sm font-semibold">원본 중복키(상위 20개)</div>
            <div className="max-h-56 overflow-auto rounded-lg border">
              <table className="w-full text-sm">
                <thead className="sticky top-0 bg-white">
                  <tr>
                    <th className="border-b p-2 text-left">상품주문번호</th>
                    <th className="border-b p-2 text-left">중복 개수</th>
                  </tr>
                </thead>
                <tbody>
                  {localResult.duplicates.slice(0, 20).map((d) => (
                    <tr key={d.key}>
                      <td className="border-b p-2">{d.key}</td>
                      <td className="border-b p-2">{d.count}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        ) : null}
      </section>
    </main>
  );
}
