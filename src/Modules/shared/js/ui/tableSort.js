export function setupSortableTable({ table, columns = [], initialSort, onSortChange }) {
  if (!table || !table.tHead) {
    return {
      getSortState: () => ({ key: "", direction: "asc" }),
      setSort: () => {}
    };
  }

  const sortableColumns = columns.filter((col) => Number.isInteger(col.index) && String(col.key || "").trim());
  if (!sortableColumns.length) {
    return {
      getSortState: () => ({ key: "", direction: "asc" }),
      setSort: () => {}
    };
  }

  const keyToColumn = new Map(sortableColumns.map((col) => [String(col.key), col]));
  const allHeaders = Array.from(table.tHead.querySelectorAll("th"));

  const initialKey =
    (initialSort && keyToColumn.has(String(initialSort.key || "")) && String(initialSort.key)) ||
    String(sortableColumns[0].key);
  const initialDirection = String(initialSort?.direction || "asc").toLowerCase() === "desc" ? "desc" : "asc";
  const state = { key: initialKey, direction: initialDirection };

  function applyHeaderState() {
    allHeaders.forEach((th, index) => {
      const column = sortableColumns.find((col) => col.index === index);
      if (!column) {
        th.classList.remove("tableSortHeader", "isSortedAsc", "isSortedDesc");
        th.removeAttribute("role");
        th.removeAttribute("tabindex");
        th.removeAttribute("aria-sort");
        th.removeAttribute("data-sort-key");
        return;
      }

      const key = String(column.key);
      th.classList.add("tableSortHeader");
      th.setAttribute("role", "button");
      th.setAttribute("tabindex", "0");
      th.dataset.sortKey = key;

      const isActive = state.key === key;
      th.classList.toggle("isSortedAsc", isActive && state.direction === "asc");
      th.classList.toggle("isSortedDesc", isActive && state.direction === "desc");
      th.setAttribute("aria-sort", isActive ? (state.direction === "asc" ? "ascending" : "descending") : "none");
    });
  }

  function setSort(key, direction, { silent = false } = {}) {
    const normalizedKey = String(key || "");
    if (!keyToColumn.has(normalizedKey)) return;
    const normalizedDirection = String(direction || "asc").toLowerCase() === "desc" ? "desc" : "asc";
    state.key = normalizedKey;
    state.direction = normalizedDirection;
    applyHeaderState();
    if (!silent && typeof onSortChange === "function") onSortChange({ ...state });
  }

  function toggleSort(key) {
    const normalizedKey = String(key || "");
    if (!keyToColumn.has(normalizedKey)) return;
    const column = keyToColumn.get(normalizedKey);
    const defaultDirection = String(column?.defaultDirection || "asc").toLowerCase() === "desc" ? "desc" : "asc";
    const nextDirection = state.key === normalizedKey ? (state.direction === "asc" ? "desc" : "asc") : defaultDirection;
    setSort(normalizedKey, nextDirection);
  }

  const onHeaderClick = (event) => {
    const th = event.target?.closest?.("th.tableSortHeader");
    if (!th || !table.contains(th)) return;
    toggleSort(th.dataset.sortKey);
  };

  const onHeaderKeydown = (event) => {
    if (event.key !== "Enter" && event.key !== " ") return;
    const th = event.target?.closest?.("th.tableSortHeader");
    if (!th || !table.contains(th)) return;
    event.preventDefault();
    toggleSort(th.dataset.sortKey);
  };

  table.tHead.addEventListener("click", onHeaderClick);
  table.tHead.addEventListener("keydown", onHeaderKeydown);
  applyHeaderState();

  return {
    getSortState: () => ({ ...state }),
    setSort
  };
}
