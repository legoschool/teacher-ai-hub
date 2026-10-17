document.querySelectorAll('[data-filter]').forEach(button => button.addEventListener('click', () => {
  document.querySelectorAll('[data-filter]').forEach(b => b.setAttribute('aria-pressed', String(b === button)));
  document.querySelectorAll('.model-card').forEach(card => card.hidden = button.dataset.filter !== '전체' && card.dataset.kind !== button.dataset.filter);
}));
window.addEventListener('hashchange', () => {
  document.querySelectorAll('video').forEach(video => video.pause());
});
