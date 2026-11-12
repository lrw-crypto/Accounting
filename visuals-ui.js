(() => {
  const original = window.ORIGINAL_EXAM_IMAGES || {};
  const lessons = window.VISUAL_LESSONS || {};
  const concepts = (window.VISUAL_CONCEPTS || []).map(visual => ({
    test: new RegExp(visual.match, 'i'),
    visual
  }));
  const escape = value => String(value ?? '').replace(/[&<>"']/g, c =>
    ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'})[c]);

  function images(files, label, description, print) {
    return files.map((file, i) => {
      const alt = i ? label + '（續頁）' : description || label;
      const caption = label + (files.length > 1 ? ' · ' + (i + 1) : '');
      return `<figure class="original-paper"><div class="original-paper-image"><img src="${escape(file)}" alt="${escape(alt)}"></div>
        ${print ? '' : `<figcaption><button type="button" class="original-zoom" aria-pressed="false" aria-label="在本頁放大${escape(caption)}">放大原圖</button></figcaption>`}
        </figure>`;
    }).join('');
  }

  function question(q, print = false) {
    const entry = original[q.id];
    if (!entry?.question) return '';
    const options = q.options.map((value, i) => 'ABCD'[i] + '. ' + (value.startsWith('IMG:') ? '圖形選項' : value));
    return images(entry.question, '原卷題目', [q.stem, ...options].join(' '), print);
  }

  function context(q, print = false) {
    const entry = original[q.id];
    const files = original[entry?.context];
    if (!files) return '';
    return `<div class="case"><b>原卷共用題組</b>${images(files, '原卷共用題組', entry.contextText, print)}</div>`;
  }

  function selectLesson(q) {
    if (lessons[q.id]) return lessons[q.id];
    const topic = q.topic || '';
    const questionText = [q.stem, ...q.options.filter(o => !o.startsWith('IMG:'))].join(' ');
    return concepts.find(({test}) => test.test(topic))?.visual
      || concepts.find(({test}) => test.test(questionText))?.visual;
  }

  function lesson(q) {
    const visual = selectLesson(q);
    if (!visual) return '';
    const blocks = visual.flow
      ? `<ol class="visual-flow">${visual.flow.map(([label, detail]) => `<li><strong>${escape(label)}</strong><span>${escape(detail)}</span></li>`).join('')}</ol>`
      : '';
    const table = visual.table
      ? `<div class="visual-table"><table><thead><tr>${visual.table[0].map(cell => `<th scope="col">${escape(cell)}</th>`).join('')}</tr></thead>
        <tbody>${visual.table.slice(1).map(row => `<tr><th scope="row">${escape(row[0])}</th>${row.slice(1).map(cell => `<td>${escape(cell)}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`
      : '';
    const graph = visual.svg ? `<img class="visual-graph" src="${escape(visual.svg)}" alt="${escape(visual.alt)}">` : '';
    return `<figure class="visual-lesson"><figcaption><b>圖解｜${escape(visual.title)}</b></figcaption>${blocks}${table}${graph}
      <p>${escape(visual.note)}</p></figure>`;
  }

  window.QuizVisuals = {hasQuestion: q => Boolean(original[q.id]?.question), question, context, lesson};

  document.addEventListener('click', event => {
    const button = event.target.closest('.original-zoom') || event.target.closest('.original-paper-image img')?.closest('.original-paper')?.querySelector('.original-zoom');
    if (!button) return;
    const figure = button.closest('.original-paper');
    const expanded = figure.classList.toggle('is-zoomed');
    button.setAttribute('aria-pressed', String(expanded));
    button.textContent = expanded ? '還原大小' : '放大原圖';
    button.setAttribute('aria-label', button.textContent);
    const viewer = figure.querySelector('.original-paper-image');
    viewer.scrollLeft = 0;
    if (expanded) viewer.setAttribute('tabindex', '0');
    else viewer.removeAttribute('tabindex');
    figure.scrollIntoView({block:'nearest'});
  });
})();
