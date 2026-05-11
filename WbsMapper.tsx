import { useRef, useState } from "react";

const API_URL = import.meta.env.VITE_WBS_API_URL ?? "http://localhost:8000";

interface Props {
  version?: string;
  projectName?: string;
}

export default function WbsMapper({ version, projectName }: Props) {
  const [downloadFile, setDownloadFile] = useState<File | null>(null);
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  const [previousFile, setPreviousFile] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const dlRef = useRef<HTMLInputElement>(null);
  const tplRef = useRef<HTMLInputElement>(null);
  const prevRef = useRef<HTMLInputElement>(null);

  const handleMap = async () => {
    if (!downloadFile || !templateFile) {
      setError("두 파일을 모두 선택해주세요.");
      return;
    }
    setError(null);
    setLoading(true);
    try {
      const body = new FormData();
      body.append("download", downloadFile);
      body.append("template", templateFile);
      if (previousFile) body.append("previous", previousFile);

      const params = new URLSearchParams();
      if (version) params.set("version", version);
      if (projectName) params.set("project", projectName);
      const qs = params.toString();
      const res = await fetch(`${API_URL}/map-wbs${qs ? "?" + qs : ""}`, { method: "POST", body });

      if (!res.ok) {
        const msg = await res.json().catch(() => ({ detail: res.statusText }));
        throw new Error(msg.detail ?? "서버 오류");
      }

      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "output_WBS.xlsx";
      a.click();
      URL.revokeObjectURL(url);
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="flex flex-col gap-4 p-6 max-w-md mx-auto">
      <h2 className="text-xl font-semibold">WBS 자동 매핑</h2>

      <label className="flex flex-col gap-1 text-sm">
        <span className="font-medium">① 다운로드 WBS (Lovable 내보내기)</span>
        <input
          ref={dlRef}
          type="file"
          accept=".xlsx"
          onChange={(e) => setDownloadFile(e.target.files?.[0] ?? null)}
          className="border rounded p-2"
        />
        {downloadFile && <span className="text-gray-500">{downloadFile.name}</span>}
      </label>

      <label className="flex flex-col gap-1 text-sm">
        <span className="font-medium">② 원본 WBS 템플릿</span>
        <input
          ref={tplRef}
          type="file"
          accept=".xlsx"
          onChange={(e) => setTemplateFile(e.target.files?.[0] ?? null)}
          className="border rounded p-2"
        />
        {templateFile && <span className="text-gray-500">{templateFile.name}</span>}
      </label>

      <label className="flex flex-col gap-1 text-sm">
        <span className="font-medium">③ 이전 출력 WBS <span className="text-gray-400">(선택 — 버전 히스토리 누적)</span></span>
        <input
          ref={prevRef}
          type="file"
          accept=".xlsx"
          onChange={(e) => setPreviousFile(e.target.files?.[0] ?? null)}
          className="border rounded p-2"
        />
        {previousFile && <span className="text-gray-500">{previousFile.name}</span>}
      </label>

      {error && <p className="text-red-500 text-sm">{error}</p>}

      <button
        onClick={handleMap}
        disabled={loading || !downloadFile || !templateFile}
        className="bg-blue-600 text-white py-2 rounded disabled:opacity-50"
      >
        {loading ? "처리 중..." : "매핑 후 다운로드"}
      </button>
    </div>
  );
}
