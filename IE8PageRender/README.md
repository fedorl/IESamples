# IE 8 page rendering sample

#### IE8 rendering problem

In IE8 renderer there is a rendering bug that prevents creating pages rasterization in an up-scaled context. Draw() will be clipped at pixel dimensions of actual display area (the visible rectangle you see is the original display area in the scale of rendering context). 

So, if there is a scaled target that is larger than actual size, while scaled, it will be clipped to the size in scaled pixels anyway (so it has much less content than original rectangle). 

If it doesn't clip for somebody on genuine IE8, then there probably are remainders of later IE in the system or there is other non-scratch setup, system update or update. 

##### Workaround possibilities

It is possible to workaround, though workarounds are a bit nasty. 

First, it is possible to workaround with IViewObject interface. But, because there is arbitrary scaling involved and accessible source rectangle is very small, this solution has some complexities. 

Instead, it can be rendered through another, now outdated API: *IHTMLElementRender*. It allows to render the page with DrawToDC into arbitrary context. Unfortunately, it is not so simple as it may appear and goes beyond just providing device context. 

First, there is the bug of clipping. It can be handled easier than IViewObject, because clipping occurs at large values beyond screen dimensions.  Second, when using device context transformations it either will not work or will mess the rendered html, so one   can't actually rely on scale or translate. Both problems require relatively non-trivial handling and complicate one another.

#### The solution 

Below is somewhat non-optimal but working on most simple pages solution. In general it is possible to achieve a perfect and more efficient solution. Obviously it is IE8 only, so one needs to check browser version and execute different handlers for IE8 or IE9 or higher. Some ideas here can apply to improve newer browsers rendering as well.

There are two interrelated workarounds here. 

##### Up-scaling 

First is: how do we upscale the vector content to the printer quality if we can't transform? The workaround here is to render to a context compatible with printer dc. What will happen is that content will be rendered at printer DPI. Note it will not fit exactly printer width, it will scale to printerDPI/screenDPI. 

Later, on rasterization, we downscale to fit the printer width. We render initially to EMF, so there is no much of quality loss that will not occur on printer itself anyway. If we need higher quality (I doubt it), there are two possibilities - modify the solution to render for target width (this is not trivial) or work with resulting emf instead of bitmap and let printer to make the downscale fit. 

Another note is that one should take not just printer width, but also possible non-printable margins that need to be queried from printer and handled appropriatelym. So it may re-scaled by printer even if one provides bitmap in exact printer dimensions. But again, I doubt this resolution disparity will make any difference for a real project. 

##### Clipping

Second to overcome is clipping. To do that, we render content in chunks less than window size which are not clipped by the renderer. After rendering a chunk, we change the scroll position of the document and render next chunk to the appropriate position in the target DC. This can be optimized to use larger chunks e.g. nearest DPI multiple to 1024 with window resize, but I didn't implement that (it is just speed optimization). If one doesn't make this optimization, there is a need to ensure that minimum browser window size is not too small.

Note, doing this scroll on arbitrary fractional scale will be an approximation and is not so simple to implement in a generic case. But with regular printer and screen we can make integer chunk steps in multiplies of DPI, e.g. if screen is 96 DPI and printer is 600DPI we make steps in the same multiple of 96 and 600 on each context and everything is much simpler. However, the remainder from top or bottom after processing all whole chunks will not be in DPI multiplies so we can't scroll as easily there. 

In genreal, we can approximate scroll position in printer space and hope there will be no misfit between chunks. Instead, what I did, is to append an absolutely positioned div with chunk size at right bottom of the page. 

Note, this can interfere with some pages and change the layout (probably not the case with simple reports). If that is a problem, one needs to add remainder handling after loops, instead of adding an element. Most simple solution in that case is still to pad with div but not with the full chunk size, just making content width multiple of screen DPI. 

Even simpler solution, I thought after, is just to resize window to the nearest multiple of DPI and take this window size as a chunk size. One can try that instead of div, this will simplify the code and fix pages that may interfere with the injected div.

## The code

This is just a sample. 

- No error handling. No checks for every COM and API call, etc.

- No code style, just quick and dirty.

- Not sure all acuired resources are released as needed

- One **have to** disable page borders on browser control for this sample to work (if ther is a need for borders around browser, just render them separately, built-in are not consistent anyway). On IE8 this is not so trivial but there are many answers on the web. Anyway, the patch for this is included in sample solution project. It is possible to  render with borders as well and exclude them, but this is unnesesary complication for a problem that has simple solution.

Sample code renders a google page and saves in c:\temp\test.emf + c:\temp\test.bmp.
