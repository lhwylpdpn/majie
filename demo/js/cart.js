(function () {
  var cart = JSON.parse(localStorage.getItem('demo_cart') || '[]');

  var cartEmpty = document.getElementById('cartEmpty');
  var cartList = document.getElementById('cartList');
  var cartSummary = document.getElementById('cartSummary');
  var selectAll = document.getElementById('selectAll');

  function saveCart() {
    localStorage.setItem('demo_cart', JSON.stringify(cart));
  }

  function formatPrice(n) {
    return n.toFixed(2);
  }

  function render() {
    if (cart.length === 0) {
      cartEmpty.classList.remove('hidden');
      cartList.classList.add('hidden');
      cartSummary.classList.add('hidden');
      return;
    }

    cartEmpty.classList.add('hidden');
    cartList.classList.remove('hidden');
    cartSummary.classList.remove('hidden');

    cartList.innerHTML = cart.map(function (item, index) {
      return (
        '<div class="cart-item" data-index="' + index + '">' +
          '<label><input type="checkbox" class="item-check" checked data-index="' + index + '"></label>' +
          '<div class="cart-item-img" data-name="' + item.name.charAt(0) + '"></div>' +
          '<div class="cart-item-info">' +
            '<h3>' + item.name + '</h3>' +
            '<span class="price-current">¥' + item.price + '</span>' +
          '</div>' +
          '<div class="cart-item-qty">' +
            '<button type="button" class="qty-btn" data-action="minus" data-index="' + index + '">-</button>' +
            '<span>' + item.qty + '</span>' +
            '<button type="button" class="qty-btn" data-action="plus" data-index="' + index + '">+</button>' +
          '</div>' +
          '<div class="cart-item-subtotal">¥' + formatPrice(item.price * item.qty) + '</div>' +
          '<button type="button" class="btn-text btn-remove" data-index="' + index + '">删除</button>' +
        '</div>'
      );
    }).join('');

    updateSummary();
  }

  function updateSummary() {
    var checks = cartList.querySelectorAll('.item-check:checked');
    var total = 0;
    var count = 0;
    checks.forEach(function (chk) {
      var idx = parseInt(chk.dataset.index, 10);
      total += cart[idx].price * cart[idx].qty;
      count += cart[idx].qty;
    });
    document.getElementById('selectedCount').textContent = count;
    document.getElementById('totalPrice').textContent = formatPrice(total);
  }

  cartList.addEventListener('click', function (e) {
    var idx = parseInt(e.target.dataset.index, 10);

    if (e.target.classList.contains('qty-btn')) {
      if (e.target.dataset.action === 'minus' && cart[idx].qty > 1) {
        cart[idx].qty -= 1;
      } else if (e.target.dataset.action === 'plus') {
        cart[idx].qty += 1;
      }
      saveCart();
      render();
    }

    if (e.target.classList.contains('btn-remove')) {
      cart.splice(idx, 1);
      saveCart();
      render();
    }
  });

  cartList.addEventListener('change', function (e) {
    if (e.target.classList.contains('item-check')) {
      updateSummary();
    }
  });

  selectAll.addEventListener('change', function () {
    cartList.querySelectorAll('.item-check').forEach(function (chk) {
      chk.checked = selectAll.checked;
    });
    updateSummary();
  });

  document.getElementById('clearCart').addEventListener('click', function () {
    cart = [];
    saveCart();
    render();
  });

  document.getElementById('checkoutBtn').addEventListener('click', function () {
    var count = parseInt(document.getElementById('selectedCount').textContent, 10);
    if (count === 0) {
      alert('请先选择商品');
      return;
    }
    alert('演示模式：结算流程待接入（共 ' + count + ' 件，¥' + document.getElementById('totalPrice').textContent + '）');
  });

  render();
})();
