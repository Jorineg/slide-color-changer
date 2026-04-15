import Pickr from '@simonwep/pickr';
import '@simonwep/pickr/dist/themes/nano.min.css';

let activePickr = null;

/**
 * Open a color picker anchored to a swatch element.
 * @param {HTMLElement} anchor - The swatch element to attach to
 * @param {string} currentHex - Current hex color (no #)
 * @param {(hex: string) => void} onChange - Called with new hex (no #) on change
 */
export function openColorPicker(anchor, currentHex, onChange) {
  if (activePickr) {
    activePickr.destroyAndRemove();
    activePickr = null;
  }

  const pickr = Pickr.create({
    el: anchor,
    theme: 'nano',
    default: `#${currentHex}`,
    appClass: 'pickr-popup',
    useAsButton: true,
    position: 'left-middle',
    components: {
      preview: true,
      opacity: false,
      hue: true,
      interaction: {
        hex: true,
        input: true,
        save: true,
      },
    },
  });

  pickr.on('save', (color) => {
    if (color) {
      const hex = color.toHEXA().toString().replace('#', '').substring(0, 6).toUpperCase();
      onChange(hex);
    }
    pickr.hide();
  });

  pickr.on('hide', () => {
    setTimeout(() => {
      pickr.destroyAndRemove();
      if (activePickr === pickr) activePickr = null;
    }, 100);
  });

  activePickr = pickr;
  pickr.show();
}

/**
 * Use the native EyeDropper API if available.
 * @returns {Promise<string|null>} Hex color (no #, uppercase) or null
 */
export async function useEyeDropper() {
  if (!('EyeDropper' in window)) return null;
  try {
    const dropper = new EyeDropper();
    const result = await dropper.open();
    return result.sRGBHex.replace('#', '').toUpperCase();
  } catch {
    return null;
  }
}
