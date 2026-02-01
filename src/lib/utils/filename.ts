export const todayKst = () => {
  // 브라우저 현지 시간 기준이지만, 날짜 포맷만 쓰므로 충분
  const d = new Date();
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
};

export const makeDatedFileName = (suffix: string) => `${todayKst()}_${suffix}`;
