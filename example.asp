<%@ Language=VBScript %>
<% Option Explicit %>
<!--#include file="asptar.asp"-->
<%
Dim objTar

Set objTar = New Tarball

objTar.AddMemoryFile "mum.txt","Hello mum!"
objTar.AddMemoryFile "dad.txt","Hello dad!"

objTar.WriteTar
%>