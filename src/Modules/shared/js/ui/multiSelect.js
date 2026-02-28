export function buildCheckboxMenu(container, values, inputName, placeholder) {
  container.innerHTML = "";

  const searchWrap = document.createElement("div");
  searchWrap.className = "multiSearchWrap";
  const searchInput = document.createElement("input");
  searchInput.className = "multiSearchInput";
  searchInput.type = "search";
  searchInput.placeholder = placeholder;
  searchInput.autocomplete = "off";
  searchInput.spellcheck = false;
  searchWrap.appendChild(searchInput);

  const optionsWrap = document.createElement("div");
  optionsWrap.className = "multiOptionsList";
  const empty = document.createElement("div");
  empty.className = "multiDropdownEmpty";
  empty.textContent = "No matches";
  empty.hidden = true;

  for (const value of values) {
    const row = document.createElement("label");
    row.className = "multiOption";
    row.dataset.search = String(value).toLowerCase();

    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.name = inputName;
    cb.value = value;

    const text = document.createElement("span");
    text.textContent = value;

    row.appendChild(cb);
    row.appendChild(text);
    optionsWrap.appendChild(row);
  }

  searchInput.addEventListener("input", () => {
    const q = String(searchInput.value || "").trim().toLowerCase();
    let visibleCount = 0;
    for (const row of optionsWrap.querySelectorAll(".multiOption")) {
      const match = !q || String(row.dataset.search || "").includes(q);
      row.hidden = !match;
      if (match) visibleCount += 1;
    }
    empty.hidden = visibleCount > 0;
  });

  container.appendChild(searchWrap);
  container.appendChild(optionsWrap);
  container.appendChild(empty);
}

export function selectedValuesFromCheckboxMenu(container) {
  return Array.from(container.querySelectorAll('input[type="checkbox"]:checked')).map((el) => el.value);
}

export function clearCheckboxMenu(container) {
  for (const cb of container.querySelectorAll('input[type="checkbox"]')) cb.checked = false;
}

export function resetCheckboxMenuSearch(container) {
  const input = container.querySelector(".multiSearchInput");
  if (!input) return;
  input.value = "";
  input.dispatchEvent(new Event("input"));
}

export function updateDropdownLabel(labelEl, values, allText, singularNoun) {
  if (!values.length) {
    labelEl.textContent = allText;
    return;
  }
  if (values.length === 1) {
    labelEl.textContent = values[0];
    return;
  }
  labelEl.textContent = `${values.length} ${singularNoun}s selected`;
}
