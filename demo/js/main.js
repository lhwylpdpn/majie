(function () {
  var cart = JSON.parse(localStorage.getItem('demo_cart') || '[]');

  function updateCartCount() {
    var count = cart.reduce(function (sum, item) { return sum + item.qty; }, 0);
    var el = document.getElementById('cartCount');
    if (el) el.textContent = count;
  }

  function saveCart() {
    localStorage.setItem('demo_cart', JSON.stringify(cart));
    updateCartCount();
  }

  function addToCart(productId) {
    var product = PRODUCTS.find(function (p) { return p.id === productId; });
    if (!product) return;

    var existing = cart.find(function (item) { return item.id === productId; });
    if (existing) {
      existing.qty += 1;
    } else {
      cart.push({ id: product.id, name: product.name, price: product.price, qty: 1 });
    }
    saveCart();
    showToast('已加入购物车：' + product.name);
  }

  function showToast(msg) {
    var toast = document.createElement('div');
    toast.className = 'toast';
    toast.textContent = msg;
    document.body.appendChild(toast);
    setTimeout(function () { toast.remove(); }, 2000);
  }

  function renderProductCard(product, showSeckill) {
    var card = document.createElement('div');
    card.className = 'product-card';
    card.innerHTML =
      '<div class="product-img" data-category="' + product.category + '">' +
        '<span class="product-tag">' + product.tag + '</span>' +
      '</div>' +
      '<div class="product-info">' +
        '<h3 class="product-name">' + product.name + '</h3>' +
        '<div class="product-price">' +
          '<span class="price-current">¥' + product.price + '</span>' +
          (showSeckill ? '<span class="price-original">¥' + product.originalPrice + '</span>' : '') +
        '</div>' +
        '<button class="btn-add-cart" data-id="' + product.id + '">加入购物车</button>' +
      '</div>';
    return card;
  }

  function initProducts() {
    var seckillEl = document.getElementById('seckillProducts');
    var recommendEl = document.getElementById('recommendProducts');
    var hotEl = document.getElementById('hotProducts');

    if (seckillEl) {
      PRODUCTS.filter(function (p) { return p.tag === '秒杀'; }).slice(0, 5).forEach(function (p) {
        seckillEl.appendChild(renderProductCard(p, true));
      });
    }

    if (recommendEl) {
      PRODUCTS.slice(0, 8).forEach(function (p) {
        recommendEl.appendChild(renderProductCard(p, false));
      });
    }

    if (hotEl) {
      PRODUCTS.filter(function (p) { return p.tag === '热销'; }).forEach(function (p) {
        hotEl.appendChild(renderProductCard(p, false));
      });
    }

    document.addEventListener('click', function (e) {
      if (e.target.classList.contains('btn-add-cart')) {
        addToCart(parseInt(e.target.dataset.id, 10));
      }
    });
  }

  function initBanner() {
    var slider = document.getElementById('bannerSlider');
    var dotsEl = document.getElementById('bannerDots');
    if (!slider) return;

    var slides = slider.querySelectorAll('.banner-slide');
    var current = 0;
    var timer = null;

    slides.forEach(function (_, i) {
      var dot = document.createElement('button');
      dot.className = 'dot' + (i === 0 ? ' active' : '');
      dot.type = 'button';
      dot.dataset.index = i;
      dotsEl.appendChild(dot);
    });

    function goTo(index) {
      slides[current].classList.remove('active');
      dotsEl.children[current].classList.remove('active');
      current = (index + slides.length) % slides.length;
      slides[current].classList.add('active');
      dotsEl.children[current].classList.add('active');
    }

    function startAuto() {
      timer = setInterval(function () { goTo(current + 1); }, 4000);
    }

    function resetAuto() {
      clearInterval(timer);
      startAuto();
    }

    document.getElementById('bannerPrev').addEventListener('click', function () {
      goTo(current - 1);
      resetAuto();
    });

    document.getElementById('bannerNext').addEventListener('click', function () {
      goTo(current + 1);
      resetAuto();
    });

    dotsEl.addEventListener('click', function (e) {
      if (e.target.classList.contains('dot')) {
        goTo(parseInt(e.target.dataset.index, 10));
        resetAuto();
      }
    });

    startAuto();
  }

  function initCountdown() {
    var el = document.getElementById('countdown');
    if (!el) return;

    var seconds = 2 * 60 * 60;
    setInterval(function () {
      seconds -= 1;
      if (seconds < 0) seconds = 2 * 60 * 60;
      var h = String(Math.floor(seconds / 3600)).padStart(2, '0');
      var m = String(Math.floor((seconds % 3600) / 60)).padStart(2, '0');
      var s = String(seconds % 60).padStart(2, '0');
      el.textContent = h + ':' + m + ':' + s;
    }, 1000);
  }

  function initSearch() {
    var btn = document.getElementById('searchBtn');
    var input = document.getElementById('searchInput');
    if (!btn || !input) return;

    btn.addEventListener('click', function () {
      var keyword = input.value.trim();
      if (!keyword) {
        showToast('请输入搜索关键词');
        return;
      }
      var matched = PRODUCTS.filter(function (p) {
        return p.name.indexOf(keyword) !== -1 || p.category.indexOf(keyword) !== -1;
      });
      showToast('找到 ' + matched.length + ' 件相关商品（演示：未跳转结果页）');
    });
  }

  updateCartCount();
  initProducts();
  initBanner();
  initCountdown();
  initSearch();
})();
